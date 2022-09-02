const ExcelJs = require('exceljs');

const filePath = "./data.csv";
const fileName = "erros.xlsx";

const workbook = new ExcelJs.Workbook();
const sheet = workbook.addWorksheet('Detalhes');
const cges = ['CGE'];
const erros = ['Erro'];
// splitedMsg.at(1).startsWith("pegar")) {

workbook.csv.readFile(filePath).then((ws) => {
  for (let i = 1; i <= ws.actualRowCount; i++) {
    for (let j = 1; j <= ws.actualColumnCount; j++) {
      const val = ws.getRow(i).getCell(j);
      const splitedMsg = val.toString().split(" ");
      if (splitedMsg.at(0) === "Fundo") {
        cges.push(splitedMsg.at(1))
        erros.push(val.toString())
      }
    }
  }

  const columnCges = sheet.getColumn(1);
  columnCges.values = cges;

  const columnErros = sheet.getColumn(2);
  columnErros.values = erros;

  workbook.xlsx.writeFile(fileName)
    .then(() => {
      console.log('Arquivo processado');
    })
    .catch(err => {
      console.log(err.message);
    });
})
