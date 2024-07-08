import React, { useEffect, useState } from "react";
import ExcelJS from "exceljs";

export const getExcelColumnLabel = (index) => {
  let label = "";
  while (index >= 0) {
    label = String.fromCharCode((index % 26) + 65) + label;
    index = Math.floor(index / 26) - 1;
  }
  return label;
};

// export const getExcelColumnNumber = (label) => {
//   let columnNumber = 0;
//   for (let i = 0; i < label.length; i++) {
//     columnNumber = columnNumber * 26 + (label.charCodeAt(i) - 64);
//   }
//   return columnNumber;
// };

// export const convertMergedCells = (mergedCells) => {
//   const result = mergedCells.map((range) => {
//     const [startCell, endCell] = range.split(":");
//     const startColumn = startCell.match(/[A-Z]+/)[0];
//     const startRow = parseInt(startCell.match(/\d+/)[0], 10);
//     const endColumn = endCell.match(/[A-Z]+/)[0];
//     const endRow = parseInt(endCell.match(/\d+/)[0], 10);

//     const startColumnNumber = getExcelColumnNumber(startColumn);
//     const endColumnNumber = getExcelColumnNumber(endColumn);

//     const convertedRange = `${startColumnNumber}${startRow}:${endColumnNumber}${endRow}`;
//     return convertedRange;
//   });

//   return result;
// };

const Optimize = () => {
  const [data, setData] = useState([]);

  const names = [
    { name: "Tháng 1", cell: "D1" },
    { name: "Năm trước", cell: "D2" },
    { name: "Quỹ lương BHXH, BHYT", cell: "D4" },
    { name: "Quỹ lương BHTN", cell: "E4" },
    { name: "NLĐ (10,5%)", cell: "F3" },
    { name: "NSDLĐ (23,5%)", cell: "J3" },
    { name: "BHXH (8%)", cell: "F4" },
    { name: "BHYT (1,5%)", cell: "G4" },
    { name: "BHTN (1%)", cell: "H4" },
    { name: "Tổng NLĐ", cell: "I4" },
    { name: "BHXH (17,5%)", cell: "J4" },
    { name: "BHYT (3%)", cell: "K4" },
    { name: "BHTN (1%)", cell: "L4" },
    { name: "Tổng BHBB (21,5%)", cell: "M4" },
    { name: "KPCĐ (2%)", cell: "N4" },
    { name: "Tổng NSDLĐ", cell: "O4" },
  ];

  useEffect(() => {
    fetch("https://dummyjson.com/products")
      .then((res) => res.json())
      .then((data) => {
        // Assuming each product has a `categories` array for the 11 category columns
        const updatedData = data.products.map((product) => ({
          ...product,
          categories: Array.from(
            { length: 11 },
            (_, i) => `${product.category} ${i + 1}`
          ), // Example category data
        }));
        setData(updatedData);
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
      { header: "TT", key: "id", width: 6 },
      { header: "Đối tượng", key: "title", width: 26 },
      { header: "Nguồn", key: "brand", width: 24 },
      ...Array.from({ length: 12 }, (_, i) => ({
        header: `Category ${i + 1}`,
        key: `category${i + 1}`,
        width: 15,
      })),
      ...Array.from({ length: 12 }, (_, i) => ({
        header: `Category ${i + 1}`,
        key: `category${i + 1}`,
        width: 15,
      })),
      { header: "Giá", key: "price", width: 10 },
    ];

    sheet.columns = columns;

    // Set styles for header
    for (let i = 1; i <= 4; i++) {
      columns.forEach((col, index) => {
        const cell = sheet.getCell(i, index + 1);

        cell.value = col.header;
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin", color: { argb: "000000" } },
          left: { style: "thin", color: { argb: "000000" } },
          bottom: { style: "thin", color: { argb: "000000" } },
          right: { style: "thin", color: { argb: "000000" } },
        };
        cell.font = {
          name: "Roboto",
          family: 4,
          size: 10,
          bold: true,
        };
        cell.fill = cell._column._key.includes("category") && {
          type: "pattern",
          pattern: "darkHorizontal",
          fgColor: { argb: "FFFF00" },
        };
      });
    }

    // Merge header cells for columns A, B, and C
    sheet.mergeCells("A1:A4");
    sheet.mergeCells("B1:B4");
    sheet.mergeCells("C1:C4");
    sheet.mergeCells("P1:P4");

    sheet.mergeCells("D1:O1"); // Merge cells for the 12 category columns header
    sheet.mergeCells("D2:O2");
    sheet.mergeCells("D3:D4");
    sheet.mergeCells("E3:E4");
    sheet.mergeCells("F3:I3");
    sheet.mergeCells("J3:O3");

    // Assign names to corresponding cells
    const reNameCell = (name, cellIndex) => {
      const cell = sheet.getCell(cellIndex);
      cell.value = name;
    };

    names?.forEach((obj) => {
      console.log(obj);
      reNameCell(obj.name, obj.cell);
    });

    // Add data to the worksheet
    data.forEach((product, index) => {
      const row = sheet.getRow(index + 5); // Start from row 5
      row.values = [
        product.id,
        product.title,
        product.brand,
        ...[...product.categories, `${product.categories[1]}2`], // Spread the categories across 12 columns
        ...[...product.categories, `${product.categories[1]}2`], // Spread the categories across 12 columns
        product.price, // Price column is separate
      ];
      row.alignment = { vertical: "middle", horizontal: "center" };
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
            <th colSpan={10}>Date Range</th>
          </tr>
          <tr>
            <th>A</th>
            <th>B</th>
            <th>C</th>
            <th colSpan={11}>Category</th>
            <th>D</th>
            <th>E</th>
            <th>F</th>
            <th>G</th>
            <th>H</th>
            <th>I</th>
            <th>J</th>
            <th>K</th>
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              <td>{row.id}</td>
              <td>{row.title}</td>
              <td>{row.brand}</td>
              {row.categories.map((category, i) => (
                <td key={i}>{category}</td>
              ))}
              <td>{row.price}</td>
              <td>A2</td>
              <td>B2</td>
              <td>C2</td>
              <td>D2</td>
              <td>E2</td>
              <td>F2</td>
              <td>G2</td>
            </tr>
          ))}
        </tbody>
      </table>
      <button onClick={exportFile}>Export</button>
    </>
  );
};

export default Optimize;
