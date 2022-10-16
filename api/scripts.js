import readXlsxFile from 'read-excel-file'

const btnConvert = document.getElementById('convert')
const inputFile = document.getElementById('excel-file')

btnConvert.addEventListener('click', () => {
    console.log('Teste')
    // readXlsxFile(inputFile.files[0]).then(rows => {
    //     console.log(rows[0])
    // })
})