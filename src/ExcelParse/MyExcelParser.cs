using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParse
{
    class MyExcelParser
    {

        public bool ParseExcelFileToDataSet(ref ExcelIntegration excel, ref DataSet inputRecords, string tableName, string filePath, string workSheetName = @"page")
        {
            Debug.Assert(inputRecords != null);
            Debug.Assert(excel != null);
            Debug.Assert(string.IsNullOrEmpty(tableName) == false);

            Debug.WriteLine($"Parsing excel report and worksheet: {filePath} : {workSheetName}");

            if ((excel == null) || (inputRecords == null))
                throw new Exception("null argument");

            if ((string.IsNullOrEmpty(filePath)) || (string.IsNullOrEmpty(workSheetName)))
                throw new Exception("Please specify a valid data-file path and worksheet name");

            if ((inputRecords.Tables.Contains(tableName)) && (inputRecords.Tables[tableName].Rows.Count > 0))
                inputRecords.Tables[tableName].Clear();

            if (excel.GetWorksheetFromExcel(filePath, workSheetName, ref inputRecords, tableName) == false)
                throw new Exception($"No records available within the specified file, {filePath}, or worksheet, {workSheetName}.\n");

            if (!(inputRecords.Tables[tableName].Rows.Count > 0))
                throw new Exception($"No records imported from Cognos excel report file {filePath}.");

            return true;
        }


        public bool ExportDataSetToExcelFile(ref ExcelIntegration excel, ref DataSet outputRecords, string filePath)
        {
            Debug.Assert(outputRecords != null);
            Debug.Assert(excel != null);
            Debug.Assert(!string.IsNullOrEmpty(filePath));
            Debug.WriteLine($"Exporting to Excel file: {filePath}");

            if ((outputRecords == null) || (excel == null) || (string.IsNullOrEmpty(filePath)))
                throw new Exception($"ExcelParser::ExportToExcelFile : null argument exception.");

            excel.SaveDataSetToExcelFile(outputRecords, filePath);
            return true;
        }
    }
}
