import ExcelJS from 'exceljs';
import {dfe} from './dfe.js'

const workbook = new ExcelJS.Workbook()

const sheet = workbook.addWorksheet('Nfe')

sheet.columns = [
    {header: 'Type', key: 'type'},
    {header: 'Event', key: 'event'},
    {header: 'CNPJ', key: 'cnpj'},
    {header: '', key: ''},
    {header: 'ICMS total', key: 'vltot'},
]

sheet.addRow({
    type: dfe.dfe,
    event: dfe.event,
    cnpj: dfe.data.emit.CNPJ
})

sheet.getRow(1).font = {
    bold: true,
    color: {argb: 'FFCCCCCC'}
}

sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: {argb: 'FF000000'}
}

sheet.workbook.xlsx.writeFile('test.xlsx')