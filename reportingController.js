const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { S3, PutObjectCommand } = require('@aws-sdk/client-s3');

const dataCaptureController = require('../controllers/dataCaptureController');

const config = {
    region: 'ap-south-1', // Specify the desired AWS region (update this as needed)
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
};

const s3 = new S3(config);

// The name of the bucket that you have created
const BUCKET_NAME = 'envsagereports';

let userEmail;
const getReport = async (req, res) => {
    try {
        userEmail = req.body.userEmail;
        // Send a POST request to the URL
        const response = await dataCaptureController.getReportData(userEmail);

        // Check if the request was successful
        if (response !== null) {
            // Parse the JSON response
            const jsonData = response;
            const filePath = await generateExcelSheet(jsonData);

            // Upload the file to S3
            const fileContent = fs.readFileSync(filePath);
            const params = {
                Bucket: BUCKET_NAME,
                Key: path.basename(filePath),
                Body: fileContent
            };

            await s3.send(new PutObjectCommand(params));
            console.log(`File uploaded successfully to S3.`);

            // Send the file as a download
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename=${path.basename(filePath)}`);
            res.download(filePath, path.basename(filePath), (err) => {
                if (err) {
                    console.error("Error sending file:", err);
                } else {
                    // Remove the file from the local system after successful download
                    fs.unlink(filePath, (err) => {
                        if (err) {
                            console.error("Error deleting file:", err);
                        } else {
                            console.log("File deleted successfully from local system.");
                        }
                    });
                }
            });
        } else {
            console.error("Error " + response);
            res.status(500).send("Error occurred while fetching data.");
        }
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json(error);
    }
};



// Function to parse JSON data and generate Excel sheet
const generateExcelSheet = async (data) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('KPI_Report');

    sheet.columns = [
        { header: 'USER EMAIL', key: 'userEmail', width: 20 },
        { header: 'FREQUENCY', key: 'time', width: 20 },
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
    
    // Save the Excel file
    const currentDate = new Date().toISOString().split('T')[0].replace(/-/g, '');
    const filePath = `Report_${userEmail.replace("@", "_")}_${currentDate}_.xlsx`;
    const resolvedFilePath = path.join(__dirname, filePath);
    await workbook.xlsx.writeFile(resolvedFilePath);
    console.log("Excel sheet generated successfully.");
    return resolvedFilePath;
};

// Function to check if a given input is a valid table row
const is_valid_table_row = (rowData) => {
    // Check if the key "No_of_rows_columns" exists and its value is greater than 0
    return rowData.hasOwnProperty("No_of_rows_columns") && !rowData.hasOwnProperty("Field_button");
};

// Function to format KPIInput
const formatKpiInput = (kpiInput) => {
    if (typeof kpiInput === 'object') {
        return Object.values(kpiInput).join(':');
    }
    return String(kpiInput);
};
const readFile = async (filePath) => {
    const fs = require('fs').promises;
    return await fs.readFile(filePath);
}
module.exports = { getReport };
