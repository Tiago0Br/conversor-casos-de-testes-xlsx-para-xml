const btnConvert = document.getElementById('convert')
const inputFile = document.getElementById('excel-file')
const pFileName = document.querySelector('.file-name')
const inputDestinationFile = document.getElementById('destination-file')
const inputStartRow = document.getElementById('start-row')
const inputEndRow = document.getElementById('end-row')
const inputName = document.getElementById('name')
const inputSummary = document.getElementById('summary')
const inputPreconditions = document.getElementById('preconditions')
const buttonDownload = document.getElementById('download')
const linkDownload = document.getElementById('link-download')

inputFile.addEventListener('change', () => {
    if (!inputFile.files[0].name.endsWith('.xlsx')) {
        swalAlert('error', 'Arquivo inválido', 'Adicione um arquivo de extensão XLSX!')
        inputFile.value = ''
        return;
    }

    const file = inputFile.files[0]
    const { name: fileName, size } = file
    const fileSize = (size / 1000).toFixed(2)
    const fileNameAndSize = `${fileName} - ${fileSize}KB`
    pFileName.textContent = fileNameAndSize
})

btnConvert.addEventListener('click', () => {
    if (!inputFile.files[0] || !inputFile.files[0].name.endsWith('.xlsx')) {
        swalAlert('error', 'Arquivo não adicionado', 'Adicione um arquivo de extensão XLSX!')
        return;
    }

    if (hasFieldEmpty()) {
        swalAlert('error', 'Campos em branco', 'Por favor, preencha todos os campos!')
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

        const fileName = outputFilename.includes('.xml') ? outputFilename : `${outputFilename}.xml`
        const xml = testcases.end({ prettyPrint: true })
        const file = new File([xml], fileName, {
            type: 'text/plain'
        })
        updateDownloadButton(file)
    })
})

buttonDownload.addEventListener('click', () => {
    linkDownload.click()
})

function swalAlert(type, title, text) {
    Swal.fire({
        icon: type,
        title: title,
        text: text,
        showConfirmButton: false,
        timer: 2000
    })
}

function hasFieldEmpty() {
    return (
        inputDestinationFile.value == ""
        || inputStartRow.value == ""
        || inputEndRow.value == ""
        || inputName.value == ""
        || inputSummary.value == ""
        || inputPreconditions.value == ""
    )
}

function updateDownloadButton(file) {
    linkDownload.href = window.URL.createObjectURL(file)
    linkDownload.download = file.name || 'saida.xml'
    buttonDownload.classList.remove('disabled')
    buttonDownload.removeAttribute('disabled')

    swalAlert('success', 'Prontinho!', 'Arquivo XML gerado com sucesso!')
}