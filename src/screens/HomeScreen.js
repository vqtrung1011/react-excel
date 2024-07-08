import React, { useEffect, useState } from "react";
import { Card, Form, Table } from "react-bootstrap";
import { getCountries, getData } from "../data/apiData";
import { ScaleLoader } from "react-spinners";
import ReactExport from "react-data-export";

const ExcelFile = ReactExport.ExcelFile;
const ExcelSheet = ReactExport.ExcelFile.ExcelSheet;
const ExcelColumn = ReactExport.ExcelFile.ExcelColumn;

const HomeScreen = () => {
  const multiDataSet = [
    {
      xSteps: 1, // Will start putting cell with 1 empty cell on left most
      ySteps: 5, //will put space of 5 rows,
      columns: [
        { title: "Headings", width: { wpx: 80 } }, //pixels width
        { title: "Text Style", width: { wch: 40 } }, //char width
        { title: "Colors", width: { wpx: 90 } },
      ],
      data: [
        [
          { value: "H1", style: { font: { sz: "24", bold: true } } },
          { value: "Bold", style: { font: { bold: true } } },
          {
            value: "Red",
            style: {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "FFFF0000", theme: "1", tint: "-0.25" },
              },
              alignment: {
                vertical: "bottom",
                horizontal: "bottom",
                readingOrder: 2,
                textRotation: 255,
              },
            },
          },
        ],
        [
          { value: "H2", style: { font: { sz: "18", bold: true } } },
          { value: "underline", style: { font: { underline: true } } },
          {
            value: "Blue",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FF0000FF" } },
            },
          },
        ],
        [
          { value: "H3", style: { font: { sz: "14", bold: true } } },
          { value: "italic", style: { font: { italic: true } } },
          {
            value: "Green",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FF00FF00" } },
            },
          },
        ],
        [
          { value: "H4", style: { font: { sz: "12", bold: true } } },
          { value: "strike", style: { font: { strike: true } } },
          {
            value: "Orange",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFF86B00" } },
            },
          },
        ],
        [
          { value: "H5", style: { font: { sz: "10.5", bold: true } } },
          { value: "outline", style: { font: { outline: true } } },
          {
            value: "Yellow",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFFFFF00" } },
            },
          },
        ],
        [
          { value: "H6", style: { font: { sz: "7.5", bold: true } } },
          { value: "shadow", style: { font: { shadow: true } } },
          {
            value: "Light Blue",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFCCEEFF" } },
            },
          },
        ],
      ],
    },
    {
      columns: [
        { title: "Headings2222222222222", width: { wpx: 80 } }, //pixels width
        { title: "Text Style", width: { wch: 40 } }, //char width
        { title: "Colors", width: { wpx: 90 } },
      ],
      data: [
        [
          { value: "H1222222", style: { font: { sz: "24", bold: true } } },
          { value: "Bold", style: { font: { bold: true } } },
          {
            value: "Red",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFFF0000" } },
            },
          },
        ],
        [
          { value: "H2", style: { font: { sz: "18", bold: true } } },
          { value: "underline", style: { font: { underline: true } } },
          {
            value: "Blue",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FF0000FF" } },
            },
          },
        ],
        [
          { value: "H3", style: { font: { sz: "14", bold: true } } },
          { value: "italic", style: { font: { italic: true } } },
          {
            value: "Green",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FF00FF00" } },
            },
          },
        ],
        [
          { value: "H4", style: { font: { sz: "12", bold: true } } },
          { value: "strike", style: { font: { strike: true } } },
          {
            value: "Orange",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFF86B00" } },
            },
          },
        ],
        [
          { value: "H5", style: { font: { sz: "10.5", bold: true } } },
          { value: "outline", style: { font: { outline: true } } },
          {
            value: "Yellow",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFFFFF00" } },
            },
          },
        ],
        [
          { value: "H6", style: { font: { sz: "7.5", bold: true } } },
          { value: "shadow", style: { font: { shadow: true } } },
          {
            value: "Light Blue",
            style: {
              fill: { patternType: "solid", fgColor: { rgb: "FFCCEEFF" } },
            },
          },
        ],
      ],
    },
  ];

  const dataSet1 = [
    {
      name: "Johson",
      amount: 30000,
      sex: "M",
      is_married: true,
    },
    {
      name: "Monika",
      amount: 355000,
      sex: "F",
      is_married: false,
    },
    {
      name: "John",
      amount: 250000,
      sex: "M",
      is_married: false,
    },
    {
      name: "Josef",
      amount: 450500,
      sex: "M",
      is_married: true,
    },
  ];

  var dataSet2 = [
    {
      name: "Johnson",
      total: 25,
      remainig: 16,
    },
    {
      name: "Josef",
      total: 25,
      remainig: 7,
    },
  ];

  console.log(multiDataSet);

  return (
    <div className="container">
      <ExcelFile>
        <ExcelSheet dataSet={multiDataSet} name="Organization" />
      </ExcelFile>

      <ExcelFile element={<button>Download Data</button>}>
        <ExcelSheet data={dataSet1} name="Employees">
          <ExcelColumn label="Name" value="name" />
          <ExcelColumn label="Wallet Money" value="amount" />
          <ExcelColumn label="Gender" value="sex" />
          <ExcelColumn
            label="Marital Status"
            value={(col) => (col.is_married ? "Married" : "Single")}
          />
        </ExcelSheet>
        <ExcelSheet data={dataSet2} name="Leaves">
          <ExcelColumn label="Name" value="name" />
          <ExcelColumn label="Total Leaves" value="total" />
          <ExcelColumn label="Remaining Leaves" value="remaining" />
        </ExcelSheet>
      </ExcelFile>
    </div>
  );
};

export default HomeScreen;
