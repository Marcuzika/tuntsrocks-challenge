const fetch = require("node-fetch");
const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("Sheet");

const url = "https://restcountries.com/v3.1/all";

fetch(url)
  .then((response) => response.json())
  .then((data) => {
    //Title
    const titleStyle = wb.createStyle({
      font: {
        color: "#4F4F4F",
        size: 16,
        bold: true,
      },
      alignment: {
        horizontal: "center",
      },
    });

    ws.cell(1, 1, 1, 4, true).string("Countries List").style(titleStyle);

    // Columns Title
    const headingStyle = wb.createStyle({
      font: {
        color: "#808080",
        size: 12,
        bold: true,
      },
    });

    const columnNames = ["Name", "Capital", "Area", "Currencies"];

    let columnIndex = 1;
    columnNames.forEach((heading) => {
      ws.cell(2, columnIndex++)
        .string(heading)
        .style(headingStyle);
    });

    //Values
    for (let i = 0; i < data.length; i++) {
      //Name
      ws.cell(3 + i, 1).string(data[i].name.common);

      //Capital
      if (data[i].capital) {
        ws.cell(3 + i, 2).string(data[i].capital);
      } else {
        ws.cell(3 + i, 2).string("-");
      }

      //Area
      if (data[i].area) {
        ws.cell(3 + i, 3)
          .number(data[i].area)
          .style({ numberFormat: "#.##0,0" });
      } else {
        ws.cell(3 + i, 3).string("-");
      }

      //Currencies
      if (data[i].currencies) {
        const keys = Object.keys(data[i].currencies);
        let allCurrencies = "";
        for (let j = 0; j < keys.length; j++) {
          allCurrencies += keys[j] + ", ";
        }
        allCurrencies = allCurrencies.slice(0, -2);
        ws.cell(3 + i, 4).string(allCurrencies);
      } else {
        ws.cell(3 + i, 4).string("-");
      }
    }

    wb.write("countries-list.xlsx");
  });
