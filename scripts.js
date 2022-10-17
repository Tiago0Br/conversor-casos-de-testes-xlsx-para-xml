const btnConvert = document.getElementById('convert')
const inputFile = document.getElementById('excel-file')
const inputDestinationFile = document.getElementById('destination-file')
const inputStartRow = document.getElementById('start-row')
const inputEndRow = document.getElementById('end-row')
const inputName = document.getElementById('name')
const inputSummary = document.getElementById('summary')
const inputPreconditions = document.getElementById('preconditions')
const button = document.getElementById('download')
const linkDownload = document.getElementById('link-download')

btnConvert.addEventListener('click', () => {
    if (!inputFile.files[0] || !inputFile.files[0].name.includes('xls')) {
        alert('Selecione um arquivo XLSX!')
        return;
    }

    if (hasFieldEmpty()) {
        alert('Preencha todos os campos!')
        return;
    }

    const outputFilename = inputDestinationFile.value
    const startRow = inputStartRow.value - 1
    const endRow = inputEndRow.value - 1
    const columnName = inputName.value - 1
    const columnSummary = inputSummary.value - 1
    const columnPreconditions = inputPreconditions.value - 1

    const testcases = xmlbuilder2.create().ele('testcases')
    readXlsxFile(inputFile.files[0]).then((rows) => {
        for (let currentRow = startRow; currentRow <= endRow; currentRow++) {
            const testcase = testcases.ele('testcase')
            testcase.att('internalid', "")
            testcase.att('name', rows[currentRow][columnName])

            testcase.ele('summary').txt(rows[currentRow][columnSummary])
            testcase.ele('preconditions').txt(rows[currentRow][columnPreconditions])
        }

        const xml = testcases.end({ prettyPrint: true })
        const file = new File([xml], outputFilename, {
            type: 'text/plain'
        })
        updateDownloadButton(file)
    })
})

function hasFieldEmpty() {
    return (
        inputDestinationFile.value == "" ||
        inputStartRow.value == "" ||
        inputEndRow.value == "" ||
        inputName.value == "" ||
        inputSummary.value == "" ||
        inputPreconditions.value == ""
    )
}

function updateDownloadButton(file) {
    linkDownload.href = window.URL.createObjectURL(file)
    linkDownload.download = file.name || 'saida.xml'
    button.classList.remove('hidden')
}