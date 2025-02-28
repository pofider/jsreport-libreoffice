const { spawn } = require('child_process')
const path = require('path')
const fs = require('fs')
const url = require('url')

const getSofficePath = async () => {
    const paths = (() => {
        switch (process.platform) {
            case 'darwin': 
                return '/Applications/LibreOffice.app/Contents/MacOS/soffice'                    
            case 'linux': 
                return  [
                '/usr/bin/libreoffice', 
                '/usr/bin/soffice', 
                '/snap/bin/libreoffice', 
                '/opt/libreoffice/program/soffice', 
                '/opt/libreoffice7.6/program/soffice'
            ]                  
            case 'win32': 
                return [
                path.join(process.env['PROGRAMFILES(X86)'], 'LIBREO~1/program/soffice.exe'),
                path.join(process.env['PROGRAMFILES(X86)'], 'LibreOffice/program/soffice.exe'),
                path.join(process.env.PROGRAMFILES, 'LibreOffice/program/soffice.exe'),
            ]
        }}
    )()

    for (const p of paths) {
        try {
            await fs.access(p)
            return p
        } catch (e) {

        }        
    }

    throw new Error('Could not find soffice binary on paths ' + paths.join(', '))
}

const sofficeConvert = (options) => { 
  const stdout = []
  const stderr = []

  const {
    callback = (() => null),
    type = 'pdf',
    outdir,
    input,
  } = options;

  const args = [
    '--convert-to', type,
    input,   
    '--headless',
    `-env:UserInstallation=${url.pathToFileURL('profile')}${options.worker}`,
    '--outdir', outdir + options.worker    
  ]

  console.log(args)
    
  const childProcess = spawn('c:\\Program Files\\LibreOffice\\program\\soffice', args);

  childProcess.stdout.on('data', (data) => {
    stdout.push(data);
  })

  childProcess.stderr.on('data', (data) => {
    stderr.push(data);
  })

  childProcess.on('close', async (code) => {
    if (stderr.length) {
      const error = new Error(Buffer.concat(stderr).toString('utf8'));    
      callback(error);
      return;
    }   
    callback()
  })

  childProcess.on('error', (err) => {
    console.error(err)
  })

  return childProcess;
}

const convert = (options) => {  
    return new Promise((resolve, reject) => {
      // Assign a fake callback that would either resolve or reject the promise
      const callback = (err, result) => (err
        ? reject(err)
        : resolve(result));

      const runOptions = {
        ...options,
        callback,
      };

      sofficeConvert(runOptions)
    })
}

async function run(i) {    
    console.time(`convert${i}`)
    await convert({
        input: 'test.docx',
        outdir: `test`,
        type: 'pdf',
        worker: i
    })
    console.timeEnd(`convert${i}`)
}

async function runMany() {
    console.time('all')
    for (let i = 0; i < 1; i++) {
        await Promise.all([run(1)])
    }
    console.timeEnd('all')
}

runMany().catch(console.error)
