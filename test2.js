const { spawn } = require('child_process')

console.time('convert')

const childProcess = spawn('python', ['unoconv.py', '-f', 'pdf', 'test.docx'])

childProcess.stdout.on('data', (data) => {
  console.log(data.toString())
})

childProcess.stderr.on('data', (data) => {
    console.log(data.toString())
})

childProcess.on('close', async (code) => {  
    console.timeEnd('convert')
})

//"c:\Program Files\LibreOffice\program\soffice" --headless -env:UserInstallation=file:///C:/work/pofider/jsreport-libreoffice/profile1 --version