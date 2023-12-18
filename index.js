const fs = require('fs');
const xl = require('excel4node');



const headingColumnNames = [
  'Código do Produto',
  'Data do evento',
  'Tipo de evento',
  'Disponíveis',
  'Reservados',
  'Referência',
  'Descrição',
  'Centro',
  'Depósito',
];


let ws;
let jsonFiles;

function list() {
  jsonFiles = fs.readdirSync('/mnt/c/json');
    console.log('arquivos encontrados', jsonFiles);
    // console.log(`🚀 ~ file: index.js:27 ~ fs.readdir ~ files:`, files);
    // console.log(`🚀 ~ file: index.js:10 ~ fs.readdir ~ jsonFiles:`, jsonFiles);
  for (const file of jsonFiles) {
    const wb = new xl.Workbook();    
    console.log(`arquivo atual: ${file}`);
    
    ws = wb.addWorksheet(`${file.split('.json')[0]}`);
    let headingColumnIndex = 1; //diz que começará na primeira linha
    headingColumnNames.forEach(heading => { //passa por todos itens do array
      // cria uma célula do tipo string para cada título
      ws.cell(1, headingColumnIndex++).string(heading);
    });
    
    let items;
    items = fs.readFileSync(`/mnt/c/json/${file}`, 'utf8');
    
    items = JSON.parse(items);
    console.log('Quantidade de itens', items.length);
    let rowIndex = 2;
    for (const item of items) {
      const collumns = [1, 2, 3, 4, 5, 6, 7, 8, 9];
      for (const collumn of collumns) {
        if (collumn === 1) {
          ws.cell(rowIndex, collumn).string(file.split('.json')[0]);
        } else if (collumn == 2) {
          ws.cell(rowIndex, collumn).string(item.createdAt || '');
        } else if (collumn === 3) {
          ws.cell(rowIndex, collumn).string(item.eventType || '');
        } else if (collumn === 4) {
          ws.cell(rowIndex, collumn).string(item.eventAvailable ? item.eventAvailable.toString() : '');
        } else if (collumn === 5) {
          ws.cell(rowIndex, collumn).string(item.eventReserved ? item.eventReserved.toString() : '');
        } else if (collumn === 6) {
          ws.cell(rowIndex, collumn).string(item.eventReference || '');
        } else if (collumn === 7) {
          ws.cell(rowIndex, collumn).string(item.eventDescription || '');
        } else if (collumn === 8) {
          ws.cell(rowIndex, collumn).string(item.plant || '');
        } else if (collumn === 9) {
          ws.cell(rowIndex, collumn).string(item.plant || '');
        }
      }
      rowIndex++;
    }
    wb.write(`${file.split('.json')[0]}.xlsx`);
    // wb.write(`estoques.xlsx`);
  }
}

list();