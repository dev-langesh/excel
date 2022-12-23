const xl = require("excel4node");
const { studentRecords } = require("./data");

const wb = new xl.Workbook();

const ws = wb.addWorksheet("Sheet 1");
const ws2 = wb.addWorksheet("Sheet 2");

const student = [
  "Reg no",
  "Name",
  "Department",
  "Year",
  "purpose",
  "In time",
  "Out time",
  "Date",
  "status",
];

const fields = Object.keys(studentRecords[0]);

var style = wb.createStyle({
  font: {
    color: "#FF0800",
    size: 12,
  },
});

student.forEach((title, i) => {
  ws.cell(1, i + 1)
    .string(title)
    .style(style);
});

let col = 2;

studentRecords.forEach((obj, i) => {
  fields.forEach((key, j) => {
    ws.cell(col, j + 1)
      .string(obj[key])
      .style(style);
  });

  col++;
});

student.forEach((title, i) => {
  ws2
    .cell(1, i + 1)
    .string(title)
    .style(style);
});

col = 2;

studentRecords.forEach((obj) => {
  fields.forEach((key, j) => {
    ws2
      .cell(col, j + 1)
      .string(obj[key])
      .style(style);
  });

  col++;
});

wb.write("./sheets/student.xlsx");
