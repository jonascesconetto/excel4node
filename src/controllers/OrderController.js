var xl = require('excel4node');
var styles = require('./styles')

module.exports = {
    async store(req, res) {
      try {
        var wb = new xl.Workbook();
        var options = {
          sheetView: {
            showGridLines: false
          }
        };
      
        var ws = wb.addWorksheet('order', options);
      
        ws.column(1).setWidth(19.92);
        ws.column(2).setWidth(64.42);
        ws.column(3).setWidth(24.42);
        ws.column(4).setWidth(14.42);
        ws.column(5).setWidth(15.42);
        ws.column(6).setWidth(15.42);
        ws.column(7).setWidth(17.42);
        ws.column(8).setWidth(9.42);
        ws.column(9).setWidth(9.42);
      
        ws.addImage({
            path: './sml.png',
            type: 'picture',
            position: {
              type: 'absoluteAnchor',
              x: '4.5mm',
              y: '4mm',
            },
          });
      
        // FIELDS
        ws.cell(1, 2, 1, 9, true).string('PEDIDO DE COMPRA - LIMPEZA').style(styles.titleStyle);
        ws.cell(5, 2, 5, 9, true).string('RELATÓRIO DE MEDIÇÃO').style(styles.titleStyle);
      
        ws.cell(2, 4, 2, 7).style(styles.headerInputStyle);
        ws.cell(4, 6, 4, 7).style(styles.headerInputStyle);
        ws.cell(3, 6, 3, 7).style(styles.headerInputStyle);
        ws.cell(6, 6, 6, 7).style(styles.headerInputStyle);
        ws.cell(7, 6, 7, 7).style(styles.headerInputStyle);
      
        ws.cell(2, 2).string('Fornecedor: ').style(styles.headerStyle);
        ws.cell(3, 2).string('Data do pedido: ').style(styles.headerStyle);
        ws.cell(4, 2).string('Embarcação: ').style(styles.headerStyle);
      
        ws.cell(6, 2).string('Emitido por:').style(styles.headerStyle);
        ws.cell(7, 2).string('Comentários:').style(styles.headerStyle);
      
        ws.cell(3, 4).string('Data de Entrega:').style(styles.headerStyle);
        ws.cell(4, 4).string('Comprador:').style(styles.headerStyle);
      
        ws.cell(6, 4).string('Nota Fiscal:').style(styles.headerStyle);
        ws.cell(7, 4).string('Data de Emissão:').style(styles.headerStyle);
      
        ws.cell(2, 8, 4, 8, true).string('PO nº:').style(styles.PORMStyle);
        ws.cell(6, 8).string('Relatório').style(styles.PORMStyle);
        ws.cell(7, 8).string('Medição').style(styles.PORMStyle);;
      
        ws.cell(8, 1).string('CÓDIGO').style(styles.productRowStyle);
        ws.cell(8, 2).string('NOME FANTASIA').style(styles.productRowStyle);
        ws.cell(8, 3).string('UM').style(styles.productRowStyle);
        ws.cell(8, 4).string('QUANTIDADE').style(styles.productRowStyle);
        ws.cell(8, 5).string('PREÇO').style(styles.productRowStyle);
        ws.cell(8, 6).string('PREÇO TOTAL').style(styles.productRowStyle);
        ws.cell(8, 7).string('QUANTIDADE ENTREGUE').style(styles.productRowStyle);
        ws.cell(8, 8, 8, 9, true).string('TOTAL FORNECIDO').style(styles.productRowStyle);
        
        var {info, product} = req.body;
        console.log(req.body);

        // HEADER INFOS
        ws.cell(2, 3).string(info.provider).style(styles.headerInputStyle);
        ws.cell(3, 3).string(info.orderDate).style(styles.headerDateStyle);
        ws.cell(4, 3).string(info.vessel).style(styles.headerInputStyle);

        ws.cell(6, 3).string(info.issuedBy).style(styles.headerInputStyle);
        ws.cell(7, 3).string(info.comments).style(styles.headerInputStyle);

        ws.cell(3, 5).string(info.deliveryDate).style(styles.headerDateStyle);
        ws.cell(4, 5).string(info.buyer).style(styles.headerInputStyle);

        ws.cell(6, 5).string(info.fiscalNote).style(styles.headerInputStyle);
        ws.cell(7, 5).string(info.issueDate).style(styles.headerDateStyle);

        ws.cell(2, 9, 4, 9, true).string(info.po).style(styles.PORMInputStyle);
        ws.cell(6, 9, 7, 9, true).string(info.measureReport).style(styles.PORMInputStyle);

        var lastRow = product.length + 9;

        for (let i = 0; i < product.length; i++) {
            
            // PRODUCTS INFO
            ws.cell(9+i, 1).string(product[i].code).style(styles.productStyle);
            ws.cell(9+i, 2).string(product[i].name).style(styles.productStyle);
            ws.cell(9+i, 3).string(product[i].um).style(styles.productStyle);
            ws.cell(9+i, 4).string(product[i].amount).style(styles.productStyle);
            ws.cell(9+i, 5).number(Number(product[i].price)).style(styles.numberStyle);
            ws.cell(9+i, 6).formula(`D${9+i} * E${9+i}`).style(styles.numberStyle);
            ws.cell(9+i, 7).style(styles.productStyle);
            ws.cell(9+i, 8, 9+i, 9, true).style(styles.productStyle);
            
        }

        ws.cell(lastRow, 1, lastRow, 5, true).string('TOTAL').style(styles.totalStyle);
        ws.cell(lastRow, 6).formula(`=SUM(F9:F${lastRow-1})`).style(styles.numberStyle);

        wb.write('Order.xlsx'); // create file in folder
        const buffer = await wb.writeToBuffer(); // byte array used to return the data to front-end to create a download link
        console.log(buffer);
        res.send(buffer); 
        return buffer;
      } catch (error) {
        console.log(error);
        return res.status(500).json({error});
      }
    }
}