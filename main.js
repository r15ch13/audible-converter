#!/usr/bin/env node
'use strict'

const _ = require('lodash')
const colors = require('colors')
const ffmpeg = require('fluent-ffmpeg')
const glob = require('glob')
const os = require('os')
const fs = require('fs')
const path = require('path')
const program = require('commander')
const sanitize = require('sanitize-filename')
const winston = require('winston')

const Promise = require('bluebird')
const ffprobe = Promise.promisify(ffmpeg.ffprobe)
const open = Promise.promisify(fs.open)
const read = Promise.promisify(fs.read, {multiArgs: true})
const pkg = require('./package.json')

const AudibleDevicesKey = 'HKLM\\SOFTWARE\\WOW6432Node\\Audible\\SWGIDMAP'
let regeditList = null
if (os.platform() === 'win32') {
  try {
    regeditList = Promise.promisify(require('regedit').list)
  } catch (e) {}
}

colors.setTheme({
  error: 'red',
  warn: 'yellow',
  info: 'green',
  verbose: 'cyan',
  debug: 'blue',
  silly: 'magenta'
})

let increaseVerbosity = (v, total) => {
  return total + 1
}

let setupWinston = () => {
  winston.level = _.findKey(winston.config.npm.levels, (o) => {
    return o === program.verbose
  })
  winston.level = winston.level || 'error'
}

let toHex = (d) => {
  // http://stackoverflow.com/a/13240395/2710739
  return ('0' + (Number(d).toString(16))).slice(-2).toUpperCase()
}

let extractBytes = (byteArray) => {
  return bytesToHex(byteArray.slice(0, 4).reverse())
}

let bytesToHex = (byteArray) => {
  return _.map(byteArray, (n) => {
    return toHex(n)
  }).join('')
}

let fetchActivationBytesFromDevices = () => {
  if (typeof regeditList !== 'function') {
    return Promise.reject(new Error(`Optional dependency \`regedit\` is not installed. Try reinstalling ${pkg.name} without ignoring \`optionalDependencies\`.`))
  }
  return regeditList(AudibleDevicesKey).then((result) => {
    let entries = _.map(result[AudibleDevicesKey].values, (n) => {
      return extractBytes(n.value)
    })

    entries = _.reject(entries, (n) => {
      return n.toUpperCase() === 'FFFFFFFF'
    })

    if (entries.length > 0) return entries
    throw new Error('Could not find any Audible Activation Bytes!')
  })
}

let fetchActivationBytes = () => {
  if (os.platform() !== 'win32') {
    return Promise.resolve(program.activationBytes)
  }

  return fetchActivationBytesFromDevices()
    .catch((err) => {
      throw err
    })
    .then((devices) => {
      let bytes = _.first(devices) || ''
      bytes = program.activationBytes ? program.activationBytes : bytes
      bytes = devices[program.device] ? devices[program.device] : bytes

      if (program.device && !devices[program.device]) {
        throw new Error(`Device Nr. ${program.device} not found! Please use the 'list' command to get your devices.`)
      }
      return bytes
    })
}

let fetchMetadata = (input) => {
  return Promise.all([
    ffprobe(input),
    fetchChecksum(input)
  ])
  .spread((result, checksum) => {
    return {
      filetype: result.format.tags.major_brand ? result.format.tags.major_brand.trim().toLowerCase() : '',
      artist: result.format.tags.artist ? result.format.tags.artist.trim() : '',
      title: result.format.tags.title ? result.format.tags.title.trim() : '',
      date: result.format.tags.date ? result.format.tags.date.trim() : '',
      duration: `${Math.floor(result.format.duration / 3600)}h${Math.floor(result.format.duration % 3600 / 60)}m${Math.floor(result.format.duration % 3600 % 60)}s`,
      durationRaw: Math.floor(result.format.duration),
      checksum: bytesToHex(checksum).toLowerCase()
    }
  })
  .then((metadata) => {
    console.log(`${metadata.artist} - ${metadata.title} [${metadata.date}] (Duration: ${metadata.duration})`)
    if (metadata.filetype !== 'aax') throw new Error('Not a valid AAX File!')
    return metadata
  })
}

let fetchChecksum = (file) => {
  let buffer = Buffer.alloc(20)
  return open(file, 'r')
  .then((fd) => {
    return read(fd, buffer, 0, buffer.length, 653) // start at absolute postion 0x28d
  })
  .catch((err) => {
    winston.error(err.message)
    winston.debug(err)
  })
  .return(buffer)
}

let currentTimemarkToPercent = (timemark, total) => {
  timemark = timemark.split(':')
  return Math.floor(((timemark[0] * 3600) + (timemark[1] * 60) + Math.floor(timemark[2])) * 100 / total)
}

let extractCoverImage = (input, output) => {
  return new Promise((resolve, reject) => {
    ffmpeg(input)
      .output(output)
      .on('end', () => {
        console.log('100%')
        resolve()
      })
      .on('error', (err) => {
        console.log('') // fix stdout
        reject(err)
      })
      .on('start', (cmd) => {
        process.stdout.write('Extracting Cover Image ... ')
        winston.debug(cmd)
      })
      .on('progress', (msg) => {
        winston.silly(msg)
      })
      .run()
  })
}

let convertAudiobook = (input, output, activationBytes, duration) => {
  return new Promise((resolve, reject) => {
    ffmpeg(input)
      .audioCodec('copy')
      .noVideo()
      .inputOptions([`-activation_bytes ${activationBytes}`])
      .output(output)
      .on('end', () => {
        console.log('') // fix stdout
        resolve()
      })
      .on('error', (err) => {
        console.log('') // fix stdout
        reject(err)
      })
      .on('start', (cmd) => {
        winston.debug(cmd)
      })
      .on('progress', (msg) => {
        process.stdout.write(`Converting Audiobook (using ${activationBytes} for decryption) ... ${currentTimemarkToPercent(msg.timemark, duration)}%` + '\r')
      })
      .run()
  })
}

let addLoopedImage = (input, output, image, duration) => {
  return new Promise((resolve, reject) => {
    ffmpeg(input)
      .input(image)
      .inputOptions(['-r 1', '-loop 1'])
      .audioCodec('copy')
      .outputOptions(['-shortest'])
      .output(output)
      .on('end', () => {
        console.log('') // fix stdout
        resolve()
      })
      .on('error', (err) => {
        console.log('') // fix stdout
        reject(err)
      })
      .on('start', (cmd) => {
        winston.debug(cmd)
      })
      .on('progress', (msg) => {
        process.stdout.write(`Adding looped cover image to Audiobook ... ${currentTimemarkToPercent(msg.timemark, duration)}%` + '\r')
      })
      .run()
  })
}

let converter = (inputFile) => {
  winston.silly(inputFile)

  let coverImage = null
  let outputFile = null
  let loopedFile = null
  let duration = 0

  return fetchMetadata(inputFile)
    .then((metadata) => {
      duration = metadata.durationRaw

      let outputDirectory = path.dirname(inputFile)
      if (program.path) {
        outputDirectory = path.resolve(program.path)
      }
      let outputFilename = sanitize(`${metadata.artist} - ${metadata.title} [${metadata.date}]`)
      if (program.output) {
        outputFilename = sanitize(path.basename(program.output, path.extname(program.output)))
      }

      coverImage = path.format({
        dir: outputDirectory,
        name: outputFilename,
        ext: '.png'
      })
      winston.silly(coverImage)
      outputFile = path.format({
        dir: outputDirectory,
        name: outputFilename,
        ext: '.m4a'
      })
      winston.silly(outputFile)
      loopedFile = path.format({
        dir: outputDirectory,
        name: outputFilename,
        ext: '.m4v'
      })
      winston.silly(loopedFile)
      return metadata
    })
    .then(() => {
      return fetchActivationBytes()
    })
    .catch((err) => {
      throw err
    })
    .then((bytes) => {
      if (!bytes) {
        throw new Error('Please provide activation bytes with -a <bytes>' + (os.platform() === 'win32' ? ' or select a device using -d <number>' : ''))
      }
      return convertAudiobook(inputFile, outputFile, bytes, duration)
    })
    .then(() => {
      return extractCoverImage(inputFile, coverImage)
    })
    .then(() => {
      if (program.loop) return addLoopedImage(outputFile, loopedFile, coverImage, duration)
    })
    .catch((err) => {
      throw err
    })
}

let globPromise = (pattern, options) => {
  return new Promise(function (resolve, reject) {
    glob(pattern, options, function (err, files) {
      return err === null ? resolve(files) : reject(err)
    })
  })
}

let main = function (inputFile) {
  setupWinston()
  globPromise(inputFile, {})
    .then((files) => {
      return Promise.reduce(files, (total, file) => {
        return converter(file).then(() => {
          console.log('')
          return ++total
        })
      }, 0)
    })
    .then((total) => {
      console.log(`Finished converting ${total > 1 ? total : 'one'} Audiobook${total > 1 ? 's' : ''}!`)
    })
    .catch((err) => {
      winston.error(err.message)
      winston.debug(err)
    })
}

program
  .version(pkg.version)
  .usage('[options] <file>')
  .option('-o, --output <filename>', 'output filename')
  .option('-p, --path <path>', 'output path')
  .option('-v, --verbose', 'output detailed information', increaseVerbosity, 0)
  .option('-a, --activation-bytes <value>', '4 byte activation secret to decrypt Audible AAX files (e.g. 1CEB00DA)', /^[A-Fa-f0-9]{8}$/i, false)
  .option('-l, --loop', 'add looped cover image to Audiobook')

if (os.platform() === 'win32') {
  program.option('-d, --device <number>', 'registered device number from which activation bytes are used')
  program
    .command('list')
    .description('list registered devices and their activation bytes')
    .action(() => {
      setupWinston()
      fetchActivationBytesFromDevices()
        .then((result) => {
          console.log('Activation bytes of registered devices:\n')
          _.each(result, (v, k) => {
            console.log(`Device ${k}: ${v}`)
          })
        })
        .catch((err) => {
          winston.error(err.message)
          winston.debug(err)
        })
    })
}

program.action(main)
program.parse(process.argv)

if (!process.argv.slice(2).length) {
  program.help()
}
