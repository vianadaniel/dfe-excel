import ExcelJS from 'exceljs';
import {dfes} from './dfe.js'

const workbook = new ExcelJS.Workbook()

const sheet = workbook.addWorksheet('Nfes')

const nfseBuilderIr = (dfes) =>{
    let nfses = []
    
    dfes.forEach( dfe => {
        const name = dfe.data.rps.prestador.razaoSocial
        const name_partner = dfe.data.rps.tomador.razaoSocial
        const cnpj = dfe.data.rps.tomador.cnpj
        const date_doc = dfe.data.rps.identificacao.dataEmissao
        const number = dfe.data.rps.identificacao.numero
        if(dfe.data.rps.servico.valores.pis?.valor > 0 && dfe.data.rps.servico.valores.cofins?.valor > 0 && 
            dfe.data.rps.servico.valores.csll?.valor > 0){

                if(dfe.data.rps.servico.valores.iss){
                    nfses.push({name, tp_imposto: 'ISS', cod_rec: '109-0', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.iss.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.iss.aliquota, tot_value: dfe.data.rps.servico.valores.iss.valorRetido})
                }

                if(dfe.data.rps.servico.valores.inss){
                    nfses.push({name, tp_imposto: 'INSS', cod_rec: '2631', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.inss.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.inss.aliquota, tot_value: dfe.data.rps.servico.valores.inss.valor})
                }

                if(dfe.data.rps.servico.valores.ir){
                    nfses.push({name, tp_imposto: 'IRRF', cod_rec: '1708', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', date_doc, number, cod_num: 'codigoNf', 
                    cal: dfe.data.rps.servico.valores.ir.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.ir.aliquota, tot_value: dfe.data.rps.servico.valores.ir.valor})
                }

            nfses.push({name, tp_imposto: 'CSRF', cod_rec: '5952', cat: 'Nota Fiscal', 
            name_partner, cnpj, date_init: 'created_at', date_doc, number, cod_num: 'codigoNf', cal: 'basecalculo', 
            aliq: '4,65', tot_value: dfe.data.rps.servico.valores.pis.valor + dfe.data.rps.servico.valores.cofins.valor
            + dfe.data.rps.servico.valores.csll.valor})
            
            }else{

                if(dfe.data.rps.servico.valores.iss){
                    nfses.push({name, tp_imposto: 'ISS', cod_rec: '109-0', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', date_doc, number, cod_num: 'codigoNf', 
                    cal: dfe.data.rps.servico.valores.iss.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.iss.aliquota, tot_value: dfe.data.rps.servico.valores.iss.valorRetido})
                }
                if(dfe.data.rps.servico.valores.inss){
                    nfses.push({name, tp_imposto: 'INSS', cod_rec: '2631', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', 
                    date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.inss.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.inss.aliquota, tot_value: dfe.data.rps.servico.valores.inss.valor})
                }
                if(dfe.data.rps.servico.valores.ir){
                    nfses.push({name, tp_imposto: 'IRRF', cod_rec: '1708', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', 
                    date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.ir.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.ir.aliquota, tot_value: dfe.data.rps.servico.valores.ir.valor})
                }
                if(dfe.data.rps.servico.valores.pis){
                    nfses.push({name, tp_imposto: 'PIS', cod_rec: '5979', cat: 'Pagamento', 
                    name_partner, cnpj, date_init: 'created_at', 
                    date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.pis.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.pis.aliquota, tot_value: dfe.data.rps.servico.valores.pis.valor})
                }
                if(dfe.data.rps.servico.valores.cofins){
                    nfses.push({name, tp_imposto: 'COFINS', cod_rec: '5960', cat: 'Pagamento', 
                    name_partner, cnpj, date_init: 'created_at', 
                    date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.cofins.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.cofins.aliquota, tot_value: dfe.data.rps.servico.valores.cofins.valor})
                }
                if(dfe.data.rps.servico.valores.csll){
                    nfses.push({name, tp_imposto: 'CSLL', cod_rec: '5987', cat: 'Nota Fiscal', 
                    name_partner, cnpj, date_init: 'created_at', 
                    date_doc, number, cod_num: 'codigoNf', cal: dfe.data.rps.servico.valores.csll.baseCalculo, 
                    aliq: dfe.data.rps.servico.valores.csll.aliquota, tot_value: dfe.data.rps.servico.valores.csll.valor})
                }
            }
    })
    return nfses
}

const dfeBuilded = nfseBuilderIr(dfes)

sheet.columns = [
    {header: 'Empresa', key: 'name'},
    {header: 'Tipo de Imposto', key: 'tp_imposto'},
    {header: 'Cód. da Receita', key: 'cod_rec'},
    {header: 'Categoria', key: 'cat'},
    {header: 'Nome Parceiro de Negócio', key: 'name_partner'},
    {header: 'CNPJ', key: 'cnpj'},
    {header: 'Data da Lançamento', key: 'date_init'},
    {header: 'Data do Documento', key: 'date_doc'},
    {header: 'Nº do Doc', key: 'number'},
    {header: 'Cód. Numérico', key: 'cod_num'},
    {header: 'Base de Cálculo', key: 'cal'},
    {header: 'Alíquota', key: 'aliq'},
    {header: 'Valor do Imposto', key: 'tot_value'},

]


dfeBuilded.forEach(each => {
    sheet.addRow(each);
});



sheet.workbook.xlsx.writeFile('testnfse.xlsx')