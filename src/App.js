import React, { useState, useEffect } from "react";
import ExcelJS from "exceljs"

function App() {
  const [excelData, setExcelData] = useState(null);
  const [showPreview, setShowPreview] = useState(false);
  const [url, setUrl] = useState(null);


const pullData = async() => {
  fetch('/data.json')
      .then(response => {
        if (!response.ok) {
          throw new Error('Network response was not ok');
        }
        return response.json();
      })
      .then(jsonData => {
        generateExcelSheet(jsonData);
      })
      .catch(error => console.error('There has been a problem with your fetch operation:', error));
}

useEffect(() => {
  pullData();
}, []);

function formatKpiInput(kpiInput) {
  if (typeof kpiInput === 'object') {
      return Object.values(kpiInput).join(':');
  }
  return String(kpiInput);
}


function is_valid_table_row(rowData) {
  // Check if the key "No_of_rows_columns" exists and its value is greater than 0
  return rowData.hasOwnProperty("No_of_rows_columns")  && !rowData.hasOwnProperty("Field_button") ;
}


const downloadExcelFile = (url) => {
  const a = document.createElement("a");
  a.href = url;
  a.download = "KPI_Report.xlsx";
  a.click();
  window.URL.revokeObjectURL(url);
}

async function generateExcelSheet(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('KPI_Report');

  sheet.columns = [
    { header: 'USER EMAIL', key: 'userEmail', width: 20 },
    { header: 'TIME', key: 'time', width: 20 },
    { header: 'KPI CODE', key: 'kpiCode', width: 15 },
    { header: 'KPI QUESTION', key: 'kpiQuestion', width: 40 },
    { header: 'KPI FORMAT', key: 'kpiFormat', width: 15 },
    { header: 'Department', key: 'department', width: 20 },
    { header: 'Operational Unit', key: 'operationalUnit', width: 20 },
    { header: 'KPI INPUT', key: 'kpiInput', width: 40 },
  ];

  // Process each entry
  let rowNumber = 2;
  data.forEach(entry => {
    const department = entry.Department;
    const operationalUnit = entry.OperationalUnit;
    const userEmail = entry.userEmail;

    // Process month frequency
    const monthFrequency = entry.monthFrequency;
    if (monthFrequency) {
      monthFrequency.forEach(monthData => {
        const month = monthData.Month;
        monthData.KPICodes.forEach(kpiCodeData => {
          const kpiCode = kpiCodeData.KPIcode;
          const kpiQuestion = kpiCodeData.KPIQuestion;
          const kpiFormat = kpiCodeData.KPIFormat;
          const kpiInput = kpiCodeData.KPIInput;

          if (Array.isArray(kpiInput)) {
            sheet.addRow({
              userEmail,
              time: month,
              kpiCode,
              kpiQuestion,
              kpiFormat,
              department,
              operationalUnit,
            });
            const columnLabels = Object.keys(kpiInput[0]);
            const sortedColumnLabels = columnLabels.sort();
            const startColumn = 8;
            let end_index = columnLabels.indexOf("No_of_rows_columns");
            const finalColumnLabels = sortedColumnLabels.slice(0, end_index);
            sheet.addRow(Array(startColumn - 1).fill("").concat(finalColumnLabels));  // Adding column headers

            kpiInput.forEach(rowData => {
              if (is_valid_table_row(rowData)) {
                const rowValues = finalColumnLabels.map(label => rowData[label] || "");
                sheet.addRow(Array(startColumn - 1).fill("").concat(rowValues));  // Adding row values
                rowNumber++;
              }
            });
            rowNumber += 2;
          } else {
            const formattedKpiInput = formatKpiInput(kpiInput);
            sheet.addRow({
              userEmail,
              time: month,
              kpiCode,
              kpiQuestion,
              kpiFormat,
              department,
              operationalUnit,
              kpiInput: formattedKpiInput
            });
          }
        });
      });
    }

    // Process annual frequency
    const annualFrequency = entry.annualFrequency;
    if (annualFrequency) {
      annualFrequency.forEach(annualData => {
        const year = annualData.Year;
        annualData.KPICodes.forEach(kpiCodeData => {
          const kpiCode = kpiCodeData.KPIcode;
          const kpiQuestion = kpiCodeData.KPIQuestion;
          const kpiFormat = kpiCodeData.KPIFormat;
          const kpiInput = kpiCodeData.KPIInput;

          if (Array.isArray(kpiInput)) {
            sheet.addRow({
              userEmail,
              time: year,
              kpiCode,
              kpiQuestion,
              kpiFormat,
              department,
              operationalUnit,
            });
            const columnLabels = Object.keys(kpiInput[0]);
            const sortedColumnLabels = columnLabels.sort();
            const startColumn = 8;
            let end_index = columnLabels.indexOf("No_of_rows_columns");
            const finalColumnLabels = sortedColumnLabels.slice(0, end_index);
            sheet.addRow(Array(startColumn - 1).fill("").concat(finalColumnLabels));  // Adding column headers

            kpiInput.forEach(rowData => {
              if (is_valid_table_row(rowData)) {
                const rowValues = finalColumnLabels.map(label => rowData[label] || "");
                sheet.addRow(Array(startColumn - 1).fill("").concat(rowValues));  // Adding row values
                rowNumber++;
              }
            });
            rowNumber += 2;
          } else {
            const formattedKpiInput = formatKpiInput(kpiInput);
            sheet.addRow({
              userEmail,
              time: year,
              kpiCode,
              kpiQuestion,
              kpiFormat,
              department,
              operationalUnit,
              kpiInput: formattedKpiInput
            });
          }
        });
      });
    }
  });


  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = window.URL.createObjectURL(blob);
  setUrl(url);
  // const a = document.createElement("a");
  // a.href = url;
  // a.download = "KPI_Report.xlsx";
  // a.click();
  // window.URL.revokeObjectURL(url);

  
  const resultArray = [];

  // console.log("data" + JSON.stringify(data))

  data.forEach(entry => {
    const userEmail = entry.userEmail;
    const department = entry.Department;
    const operationalUnit = entry.OperationalUnit;

    const monthFrequency = entry.monthFrequency;

    if (monthFrequency) {
      monthFrequency.forEach(monthData => {
        const month = monthData.Month;
      monthData.KPICodes.forEach(kpiCodeData => {
        const kpiCode = kpiCodeData.KPIcode;
      const kpiQuestion = kpiCodeData.KPIQuestion;
      const kpiFormat = kpiCodeData.KPIFormat;
      const kpiInput = kpiCodeData.KPIInput;

      const rowData = {
        'USER EMAIL': userEmail,
        'TIME': month,
        'KPIcode': kpiCode,
        'KPIQuestion': kpiQuestion,
        'KPIFormat': kpiFormat,
        // 'draftStatus': kpiCodeData.draftStatus,
        // 'captureDateTime': kpiCodeData.captureDateTime,
        "Department": department,
            "Operational Unit": operationalUnit
      };

        // Check if KPIInput is an array and KPIFormat is "table 15"
        if (Array.isArray(kpiInput) ) {
          const subTable = kpiInput.map(item => ({
            Field_1: item.Field_1,
            Field_2: item.Field_2,
            Field_3: item.Field_3,
            Field_4: item.Field_4,
            Field_5: item.Field_5,
            Field_6: item.Field_6,
            Field_7: item.Field_7,
          }));

          rowData.KPIInput = subTable;
        }
         else if (typeof kpiInput === 'object') {
          const ratioData =  String(Object.values(kpiInput).join(':'));
          rowData.KPIInput = ratioData;
      }
      else {
          rowData.KPIInput = kpiInput;
        }
          resultArray.push(rowData);
          
        });
      });
    }

    const annualFrequency = entry.annualFrequency;
    if (annualFrequency){
      annualFrequency.forEach(annualData => {
        const year = annualData.Year;
        annualData.KPICodes.forEach(kpiCodeData => {
          const kpiCode = kpiCodeData.KPIcode;
          const kpiQuestion = kpiCodeData.KPIQuestion;
          const kpiFormat = kpiCodeData.KPIFormat;
          const kpiInput = kpiCodeData.KPIInput;

          const rowData = {
            'USER EMAIL': userEmail,
            'TIME': year,
            'KPIcode': kpiCode,
            'KPIQuestion': kpiQuestion,
            'KPIFormat': kpiFormat,
            "Department": department,
            "Operational Unit": operationalUnit
            // 'draftStatus': kpiCodeData.draftStatus,
            // 'captureDateTime': kpiCodeData.captureDateTime,
          };

         

          if (Array.isArray(kpiInput)) {
            const subTable = kpiInput.map(item => ({
              Field_1: item.Field_1,
              Field_2: item.Field_2,
              Field_3: item.Field_3,
              Field_4: item.Field_4,
              Field_5: item.Field_5,
              Field_6: item.Field_6,
              Field_7: item.Field_7,
            }));
            rowData.KPIInput = subTable;
          } else if (typeof kpiInput === 'object') {
            const ratioData = String(Object.values(kpiInput).join(':'));
            rowData.KPIInput = ratioData;
          } else {
            rowData.KPIInput = kpiInput;
          }
          resultArray.push(rowData);
        });
      });
    }
  });
  // console.log("data" + JSON.stringify(resultArray))
  setExcelData(resultArray)

}


  return (
    <div className="wrapper">
    <button onClick={() =>downloadExcelFile(url)}>download</button>

      <button  onClick={() => setShowPreview(true)}>Preview</button>

      {showPreview ? <div className="viewer">
        {excelData?(
          <div className="table-responsive">
            <table className="table">

              <thead>
                <tr>
                  {Object.keys(excelData[0]).map((key)=>(
                    <th key={key}>{key}</th>
                  ))}
                </tr>
              </thead>

              <tbody>
                  {excelData.map((individualExcelData, index) => (
                    <tr key={index}>
                      {Object.keys(individualExcelData).map((key) => (
                        <td key={key}>
                          {Array.isArray(individualExcelData[key]) ? (
                            <table className="inner-table">
                              <tbody>
                                {individualExcelData[key].map((item, idx) => (
                                  <tr key={`${key}-${idx}`}>
                                    {Object.entries(item).map(([subKey, subValue]) => (
                                      <td key={`${key}-${subKey}`}>{subValue}</td>
                                    ))}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          ) : (
                            individualExcelData[key].toString()
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>

            </table>
          </div>
        ):(
          <div>No File</div>
        )}
      </div>: null}
      

    </div>
  );
}

export default App;
