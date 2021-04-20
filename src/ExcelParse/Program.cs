using CommandLine;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParse
{
    class Options
    {
        [Option('e', "ExcelFileName", Required = true, Default = null,
          HelpText = "Excel Data file.")]
        public string ExcelFileName { get; set; }

        [Option('w', "WorksheetName", Required = false, Default = "page",
          HelpText = "Excel data worksheet name. Default is 'page'.")]
        public string WorksheetName { get; set; }
    }

    class MyProgram
    {
        private string _TimeTag;                // Establish a time-stamp.
        private string _OutputFolder;           // Path to file outputs.

        private MyExcelParser _ExcelParser;     // Example Excel Parsing class

        private ExcelIntegration _ExcelApp;     // Office Interop

        private string _ExcelFilePath;          // Path to input file 
        private string _WorksheetName;          // The worksheet name, default is 'page'.

        private string _OutputFileName;         // Echo the input file and append a time-stamp to the file name.

        private DataSet _InputRecords = new DataSet("Inputs");

        static void Main(string[] args)
        {
            MyProgram myApp = new MyProgram();

            myApp._TimeTag = DateTime.Now.ToString(@"yyyy'-'MM'-'dd'-'HHmm");
            myApp._ExcelParser = null;
            myApp._ExcelApp = null;
            myApp._InputRecords = null;

            try
            {
                Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       myApp._ExcelFilePath = o.ExcelFileName;

                       if (!string.IsNullOrEmpty(o.ExcelFileName))
                       {
                           myApp._ExcelFilePath = Path.GetFullPath(o.ExcelFileName);
                           myApp._OutputFileName = myApp._ExcelFilePath.Replace(".xlsx","_") + myApp._TimeTag + ".xlsx";
                       }

                       if (!string.IsNullOrEmpty(o.WorksheetName))
                           myApp._WorksheetName = o.WorksheetName;
                   })
                   .WithNotParsed<Options>((errs) => HandleParseErrors(errs));

                myApp._InputRecords = new DataSet("Inputs");

                //
                //  Configure Excel integration.
                //
                myApp._ExcelApp = new ExcelIntegration();

                G.Display($"Importing worksheet {myApp._WorksheetName} from file {myApp._ExcelFilePath}");

                myApp._ExcelParser = new MyExcelParser();

                string tableName = myApp._WorksheetName.Trim().Replace(" ", "-");

                if (myApp._ExcelParser.ParseExcelFileToDataSet(ref myApp._ExcelApp, ref myApp._InputRecords, tableName, myApp._ExcelFilePath, myApp._WorksheetName) == false)
                    throw (new Exception($"Error importing worksheet {myApp._WorksheetName} from Excel file {myApp._ExcelFilePath}"));

                if (myApp._InputRecords.Tables[tableName].Rows.Count > 0)
                    G.Display($"Imported {myApp._InputRecords.Tables[tableName].Rows.Count} records from worksheet {myApp._WorksheetName}");

                //
                // Export the same data-set to a new Excel file.
                //
                if (myApp._ExcelParser.ExportDataSetToExcelFile(ref myApp._ExcelApp, ref myApp._InputRecords, myApp._OutputFileName) == false)
                    throw (new Exception($"Error exporting data-set to Excel {myApp._OutputFileName}"));

                G.Display($"\nSee export file: {myApp._OutputFileName}");
            }
            catch (Exception ex)
            {
                G.DisplayError(ex.Message);
            }
            finally
            {
                if (myApp._ExcelApp != null)
                    myApp._ExcelApp.Close();
                if (myApp._InputRecords != null)
                    myApp._InputRecords.Clear();
            }

            return;
        }


        private static void HandleParseErrors(IEnumerable<Error> errs)
        {
            throw (new Exception(string.Format("Error parsing command line options.")));
        }

    }  // Class MyProgram

}
