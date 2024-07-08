import React, { useEffect, useState } from "react";
import ExcelJS from "exceljs";

const NavbarScreen = () => {
  const [data, setData] = useState([]);

  useEffect(() => {
    fetch("https://dummyjson.com/products")
      .then((res) => res.json())
      .then((data) => {
        setData(data.products); // Assuming the data structure has 'products'
      });
  }, []);

  const exportFile = () => {
    const workBook = new ExcelJS.Workbook();
    const sheet = workBook.addWorksheet("My Sheet", {
      headerFooter: {
        firstHeader: "Hello Exceljs",
        firstFooter: "Hello World",
      },
    });

    // Define the columns and their properties
    const columns = [
      { header: "Id", key: "id", width: 6 },
      { header: "Title", key: "title", width: 26 },
      { header: "Brand", key: "brand", width: 24 },
      { header: "Category", key: "category", width: 22 },
      { header: "Price", key: "price", width: 10 },
    ];

    sheet.columns = columns;

    // Set headers in the first 4 rows
    for (let i = 1; i <= 4; i++) {
      columns.forEach((col, index) => {
        const cell = sheet.getCell(i, index + 1);
        cell.value = col.header;
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin", color: { argb: "000" } },
          left: { style: "thin", color: { argb: "000" } },
          bottom: { style: "thin", color: { argb: "000" } },
          right: { style: "thin", color: { argb: "000" } },
        };
        cell.fill = {
          type: "pattern",
          pattern: "darkHorizontal",
          fgColor: { argb: "FFFF00" },
        };
        cell.font = {
          name: "Roboto",
          family: 4,
          size: 10,
          bold: true,
        };
      });
    }

    // Add data to the worksheet
    data.forEach((product, index) => {
      const startRow = index * 4 + 5;
      for (let j = 0; j < 4; j++) {
        const row = sheet.getRow(startRow + j);
        row.values = [
          product.id,
          product.title,
          product.brand,
          product.category,
          product.price,
        ];
        row.alignment = { vertical: "middle", horizontal: "center" };
      }
    });

    // Apply validation and formatting to the price column
    const priceCol = sheet.getColumn("price");
    priceCol.eachCell((cell) => {
      const cellValue = cell.value;
      if (cellValue > 50 && cellValue < 1000) {
        cell.fill = {
          type: "pattern",
          pattern: "lightDown",
          fgColor: { argb: "FF0000" },
        };
      }
    });

    // Export to Excel file
    workBook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "download.xlsx";
      anchor.click();
      window.URL.revokeObjectURL(url);
    });
  };

  return (
    <>
      <table id="sampletable" className="uk-report-table table table-striped">
        <thead style={{ background: "red" }}>
          <tr>
            <th colSpan="1">Date Range</th>
            <th colSpan="5">
              <h2>Last 30 Days</h2>
            </th>
            <th colSpan="5">
              <h2>Previous 30 Days</h2>
            </th>
          </tr>
          <tr>
            <th>A</th>
            <th>B</th>
            <th>C</th>
            <th>D</th>
            <th>E</th>
            <th>A2</th>
            <th>B2</th>
            <th>C2</th>
            <th>D2</th>
            <th>E2</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
      <button onClick={exportFile}>Export</button>
    </>
  );
};

export default NavbarScreen;
