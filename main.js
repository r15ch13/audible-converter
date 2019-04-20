#!/usr/bin/env node
'use strict'

const _ = require('lodash')
const ffmpeg = require('fluent-ffmpeg')
const glob = require('glob')
const os = require('os')
const fs = require('fs')
const https = require('https')
const path = require('path')
const program = require('commander')
const sanitize = require('sanitize-filename')
const winston = require('winston')
const spawn = require('child_process').spawn

const Promise = require('bluebird')
const ffprobe = Promise.promisify(ffmpeg.ffprobe)
const open = Promise.promisify(fs.open)
const read = Promise.promisify(fs.read, {multiArgs: true})
const readFile = Promise.promisify(fs.readFile)
const pkg = require(path.join(__dirname, './package.json'))

process.env.RCRACK_PATH = process.env.RCRACK_PATH || path.resolve(path.join(__dirname, `./tools/rcrack/${os.platform()}/rcrack`))

const AudibleDevicesKey = 'HKLM\\SOFTWARE\\WOW6432Node\\Audible\\SWGIDMAP'
let regeditList = null
if (os.platform() === 'win32') {
  try {
    regeditList = Promise.promisify(require('regedit').list)
  } catch (e) {}
}

winston.addColors({
  error: 'red',
  warn: 'yellow',
  info: 'green',
  verbose: 'cyan',
  debug: 'blue',
  silly: 'magenta'
});

let increaseVerbosity = (v, total) => {
  return total + 1
}

const logger = winston.createLogger({
  level: {
    error: 0,
    warn: 1,
    info: 2,
    verbose: 3,
    debug: 4,
    silly: 5
  },
  transports: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      )
    })
  ]
});

let setupWinston = () => {
  logger.level = _.findKey(logger.levels, (o) => {
    return o === program.verbose
  })
  logger.level = logger.level || 'error'
  logger.log('debug', `RCRACK_PATH: ${process.env.RCRACK_PATH}`)
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
  if (os.platform() !== 'win32' || program.activationBytes) {
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
    logger.log('error', err.message)
    logger.log('debug', err.stack)
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
        logger.log('debug', cmd)
      })
      .on('progress', (msg) => {
        logger.log('silly', msg)
      })
      .run()
  })
}

let extractDownloadURL = (adhFile) => {
  return new Promise((resolve, reject) => {
    readFile(adhFile)
      .then((content) => {
        content = content.toString()
        let custId = content.match(/cust_id=([\w-]+[^&])/).pop()
        let productId = content.match(/product_id=([\w-]+[^&])/).pop()
        let codec = content.match(/codec=([\w-]+[^&])/).pop()
        let title = content.match(/title=([^&]+)/).pop()
        resolve({url: `https://cds.audible.de/download?product_id=${productId}&cust_id=${custId}&codec=${codec}`, title: title})
      })
      .catch((err) => {
        reject(err)
      })
  })
}

let download = (url, output) => {
  return new Promise((resolve, reject) => {
    https.get(url, (response) => {
      const statusCode = response.statusCode

      if (statusCode !== 200) {
        return reject(new Error('Download error!'))
      }

      let receivedBytes = 0
      let totalBytes = response.headers['content-length']

      response.on('data', function (chunk) {
        receivedBytes += chunk.length
        process.stdout.write(`Downloading '${output}' ... ${((receivedBytes * 100) / totalBytes).toFixed(0)}% | ${(receivedBytes / 1048576).toFixed(2)} / ${(totalBytes / 1048576).toFixed(2)} MB` + '\r')
      })

      const writeStream = fs.createWriteStream(output)
      response.pipe(writeStream)

      writeStream.on('error', (err) => {
        console.log('') // fix stdout
        reject(err)
      })
      writeStream.on('finish', () => {
        console.log('') // fix stdout
        writeStream.close(resolve)
      })
    })
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
        logger.log('debug', cmd)
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
        logger.log('debug', cmd)
      })
      .on('progress', (msg) => {
        process.stdout.write(`Adding looped cover image to Audiobook ... ${currentTimemarkToPercent(msg.timemark, duration)}%` + '\r')
      })
      .run()
  })
}

let rcrack = (checksum) => {
  return new Promise((resolve, reject) => {
    let child = spawn(process.env.RCRACK_PATH, ['tables', '-h', checksum], { cwd: path.resolve(process.env.RCRACK_PATH, '../..') })
    child.addListener('error', reject)
    child.stdout.on('data', resolve)
    child.stderr.on('data', reject)
  })
}

let lookupChecksum = (checksum) => {
  console.log(`Looking up activation bytes for checksum: ${checksum}`)
  console.log(`This might take a moment ...`)
  return rcrack(checksum)
    .then((output) => {
      logger.log('info', output.toString())
      let matches = output.toString().match(/hex:([a-fA-F0-9]{8})/)
      if (!matches) throw new Error('Activation Bytes where not found!')
      console.log('Activation Bytes found:', matches[1])
    })
}

let converter = (inputFile) => {
  logger.log('silly', inputFile)

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
      logger.log('silly', coverImage)
      outputFile = path.format({
        dir: outputDirectory,
        name: outputFilename,
        ext: '.m4a'
      })
      logger.log('silly', outputFile)
      loopedFile = path.format({
        dir: outputDirectory,
        name: outputFilename,
        ext: '.m4v'
      })
      logger.log('silly', loopedFile)
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
      logger.log('error', err.message)
      logger.log('debug', err.stack)
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
  program.option('-d, --device <number>', 'registered device number from which activation bytes are used (Windows only)')
}
program
  .command('list')
  .description('list registered devices and their activation bytes (Windows only)')
  .action(() => {
    if (os.platform() !== 'win32') {
      console.error('This command is only available on Windows')
      return
    }
    setupWinston()
    fetchActivationBytesFromDevices()
      .then((result) => {
        console.log('Activation bytes of registered devices:\n')
        _.each(result, (v, k) => {
          console.log(`Device ${k}: ${v}`)
        })
      })
      .catch((err) => {
        logger.log('error', err.message)
        logger.log('debug', err.stack)
      })
  })
program
  .command('lookup')
  .description('lookup activation bytes in RainbowTables generated by https://github.com/inAudible-NG/ (Windows/Linux only)')
  .arguments('<file|checksum>')
  .action((fileOrChecksum) => {
    if (os.platform() !== 'win32' && os.platform() !== 'linux') {
      console.error('This command is only available on Windows and Linux')
      return
    }
    setupWinston()
    if (fileOrChecksum.match(/([a-fA-F0-9]{20})/)) {
      lookupChecksum(fileOrChecksum)
        .catch((err) => {
          logger.log('error', err.message)
          logger.log('debug', err.stack)
        })
    } else {
      fetchMetadata(fileOrChecksum)
        .then((metadata) => {
          return lookupChecksum(metadata.checksum)
        })
        .catch((err) => {
          logger.log('error', err.message)
          logger.log('debug', err.stack)
        })
    }
  })
program
  .command('checksum')
  .description('show audiobooks checksum')
  .arguments('<file>')
  .action((inputFile) => {
    setupWinston()
    fetchMetadata(inputFile)
      .then((metadata) => {
        console.log(`Checksum for ${inputFile} is ${metadata.checksum}`)
      })
      .catch((err) => {
        logger.log('error', err.message)
        logger.log('debug', err.stack)
      })
  })

program
  .command('download')
  .description('download an audiobook from *.adh')
  .arguments('<file>')
  .action((inputFile) => {
    setupWinston()
    extractDownloadURL(inputFile)
    .then((result) => {
      return download(result.url, `${sanitize(result.title)}.aax`)
    })
    .then(() => {
      console.log('Download complete!')
    })
    .catch((err) => {
      logger.log('error', err.message)
      logger.log('debug', err.stack)
    })
  })

program.action(main)
program.parse(process.argv)

if (!process.argv.slice(2).length) {
  program.help()
}
