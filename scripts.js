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
const radioButtonSim = document.getElementById('radio_sim')
const radioButtonNao = document.getElementById('radio_nao')

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

    const defaultFileName = fileName.replace('.xlsx', '.xml')
    inputDestinationFile.setAttribute('placeholder', defaultFileName)

    inputEndRow.focus()
})

btnConvert.addEventListener('click', () => {
    if (!inputFile.files[0] || !inputFile.files[0].name.endsWith('.xlsx')) {
        swalAlert('error', 'Arquivo não adicionado', 'Adicione um arquivo de extensão XLSX!')
        return;
    }

    if (hasRequiredFieldEmpty()) {
        swalAlert('error', 'Campos em branco', 'Por favor, preencha o campo "Última linha"!')
        return;
    }

    const defaultFileName = inputFile.files[0].name.replace('.xlsx', '.xml')
    const outputFilename = inputDestinationFile.value || defaultFileName
    const startRow = inputStartRow.value ? inputStartRow.value - 1 : 1
    const endRow = inputEndRow.value - 1
    const columnName = inputName.value ? inputName.value - 1 : 0
    const columnSummary = inputSummary.value ? inputSummary.value - 1 : 1
    const columnPreconditions = inputPreconditions.value ? inputPreconditions.value - 1 : null

    const testcases = xmlbuilder2.create().ele('testcases')
    readXlsxFile(inputFile.files[0]).then((rows) => {
        let position = 1
        for (let currentRow = startRow; currentRow <= endRow; currentRow++) {
            const testcase = testcases.ele('testcase')
            testcase.att('internalid', "")

            if (document.querySelector('input[checked]').value === 's') {
                const testNumber = position  >= 10 ? `${position}` : `0${position}`
                testcase.att('name', `${testNumber} - ${rows[currentRow][columnName]}`)
            } else {
                testcase.att('name', rows[currentRow][columnName])
            }

            testcase.ele('summary').txt(rows[currentRow][columnSummary])
            if (columnPreconditions) {
                testcase.ele('preconditions').txt(rows[currentRow][columnPreconditions])
            }
            position++
        }

        const fileName = outputFilename.includes('.xml') ? outputFilename : `${outputFilename}.xml`
        const xml = testcases.end({ prettyPrint: true })
        const file = new File([xml], fileName, {
            type: 'text/plain'
        })
        updateDownloadButton(file)

        const quant = position-1
        swalAlert('success', 'Sucesso!', `Foram gerados ${quant} casos de testes.`)
    })
})

radioButtonSim.addEventListener('click', () => {
    radioButtonSim.setAttribute('checked', true)
    radioButtonNao.removeAttribute('checked')
})

radioButtonNao.addEventListener('click', () => {
    radioButtonNao.setAttribute('checked', true)
    radioButtonSim.removeAttribute('checked')
})

buttonDownload.addEventListener('click', () => {
    linkDownload.click()

    window.location.reload()
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

function hasRequiredFieldEmpty() {
    return (
        inputEndRow.value == ""
    )
}

function updateDownloadButton(file) {
    linkDownload.href = window.URL.createObjectURL(file)
    linkDownload.download = file.name || 'saida.xml'
    buttonDownload.classList.remove('disabled')
    buttonDownload.removeAttribute('disabled')
}