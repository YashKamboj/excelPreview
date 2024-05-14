import { Button } from "@mui/material";
import React from "react";
import { FaArrowDownLong } from "react-icons/fa6";
import axios from "axios";
import * as FileSaver from "file-saver";
import { baseUrl } from "../baseurl";

const DemoDownloadBtn = () => {

  const [showPreview, setShowPreview] = useState(false);
  const [excelData, setExcelData] = useState(null);

  const handleDownload = async () => {
    try {
      const token = localStorage.getItem("token");
      const userEmail = localStorage.getItem("email");
      const headers = {
        Authorization: `${token}`,
        "Content-Type": "application/json",
      };
      const currentDate = new Date()
        .toISOString()
        .replace(/[-T:]/g, "")
        .split(".")[0];

      const fileName = `Report_${userEmail.replace(
        "@",
        "_"
      )}_${currentDate}.xlsx`;
      const res = await axios.post(
        `${baseUrl}/getReport`,
        { userEmail },
        { headers, responseType: "blob" }
      );
      const blob = new Blob([res.data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      FileSaver.saveAs(blob, fileName);
    } catch (error) {
      console.error("Error:", error);
    }
  };

  const handlePreviewData = async () => {
    try{
      
      const token = localStorage.getItem("token");
      const userEmail = localStorage.getItem("email");
      const headers = {
        Authorization: `${token}`,
        "Content-Type": "application/json",
      };
      const res = await axios.post(
        `${baseUrl}/getPreviewData`,
        { userEmail },
        { headers, responseType: "blob" }
      );

      setExcelData(res.data);
      setShowPreview(true)

    }catch (error) {
      console.error("Error:", error);
    }
  }

  return (
    <div>
      <Button
        onClick={handleDownload}
        variant="outlined"
        startIcon={<FaArrowDownLong />}
      >
        Download
      </Button>
      <Button  onClick={() => handlePreviewData()}  variant="outlined">Preview</Button>
    
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
};

export default DemoDownloadBtn;
