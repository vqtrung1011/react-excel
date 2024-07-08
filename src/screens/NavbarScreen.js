import React from "react";
import * as XLSX from "xlsx";

const NavbarScreen = () => {
  const exportFile = () => {
    console.log("a");
    var wb = XLSX.utils.table_to_book(document.getElementById("sampletable"));
    XLSX.writeFile(wb, "sample.xlsx");
  };

  return (
    <>
      <table id="sampletable" class="uk-report-table table table-striped">
        <thead>
          <tr>
            <th colSpan="1">Date Range</th>
            <th colSpan="5" style={{ background: "red" }}>
              <strong>Last 30 Days</strong>
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
