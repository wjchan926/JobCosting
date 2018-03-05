using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace JobCosting
{
    class ExcelFileReader
    {
        private  Excel.Workbook myBook = null;
        private  Excel.Application myApp = null;
        private  Excel.Worksheet mySheet = null;

        private string jobCostingPath = @"S:\JOB COSTING REFERENCE WORK SHEET.xlsm";

        public ExcelFileReader()
        {
            // Creates new instacne of excel
            myApp = new Excel.Application();

            // True to see new instance, false to hide
            myApp.Visible = true;

            // Sets workbook to path specified
            // Try to open workbook.
            try
            {
                myBook = myApp.Workbooks.Open(jobCostingPath);
            }
            catch (NullReferenceException e)
            { 
                // If file is not found
                System.Diagnostics.Debug.WriteLine(e.Message);
                throw;
            }
            catch (Exception e)
            {
                // Other problems
                System.Diagnostics.Debug.WriteLine(e.Message);
                throw;
            }

            // Sets worksheet to specified sheet.  Starts at 1 or specify sheet name as string
            mySheet = myBook.Sheets["2018"];
               
        }

    }

    //class TestExcelFileReader
    //{
    //     void Main()
    //    {
    //        ExcelFileReader excelFileReader = new JobCosting.ExcelFileReader();            

    //    }
    //}
}
