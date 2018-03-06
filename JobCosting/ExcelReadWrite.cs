using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JobCosting
{
    sealed class ExcelReadWrite
    {
        public Excel.Workbook myBook { get; private set; } = null;
        public Excel.Application myApp { get; private set; } = null;
        public Excel.Worksheet mySheet { get; private set; } = null;
        public Excel.Range range { get; private set; } = null;

        private string jobCostingPath = @"S:\JOB COSTING REFERENCE WORK SHEET.xlsm";

        /// <summary>
        /// Default contructor for ExcelFileReaer.
        /// Called with Open Job Costing Document GUI Method.
        /// </summary>
        public ExcelReadWrite()
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

        /// <summary>
        /// Closes workbook and applicaiton.  Releases Objects
        /// Called with the Save and Close GUI Method
        /// </summary>
        public void close()
        {
            myBook.Close(true, null, null);
            myApp.Quit();

            release();
        }

        public void release()
        {
            Marshal.ReleaseComObject(mySheet);
            Marshal.ReleaseComObject(myBook);
            Marshal.ReleaseComObject(myApp);
        }

        public void setRange()
        {
            // Try to get selected range of cells from excel
            try
            {
                range = myApp.Selection;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Selection is not a valid range of cells");
            }
        }               

        // Write Methods
        /// <summary>
        /// Writes job data to the excel sheet
        /// </summary>
        /// <param name="job"></param> Job to be analyzed
        public void writeJobData(SuperJob job)
        {
            // Determine which sheet is open
            mySheet = myApp.ActiveSheet;

            foreach (Excel.Range range in range.Rows)
            {
                mySheet.Cells[range.Row, ExcelColumn.salesRep] = job.salesRep;
                mySheet.Cells[range.Row, ExcelColumn.actualCost] = job.amountActualCost;
                mySheet.Cells[range.Row, ExcelColumn.actualRevenue] = job.amountActualRevenue;
                mySheet.Cells[range.Row, ExcelColumn.difference] = job.difference;
                mySheet.Cells[range.Row, ExcelColumn.grossMargin] = job.grossMargin;
                mySheet.Cells[range.Row, ExcelColumn.unitHigh] = job.unitHigh;
                mySheet.Cells[range.Row, ExcelColumn.unitMed] = job.unitMed;
                mySheet.Cells[range.Row, ExcelColumn.unitLow] = job.unitLow;
                mySheet.Cells[range.Row, ExcelColumn.unitFloor] = job.unitFloor;
                mySheet.Cells[range.Row, ExcelColumn.freight] = job.freight;
                mySheet.Cells[range.Row, ExcelColumn.marlinFreight] = job.marlinFreight;
                mySheet.Cells[range.Row, ExcelColumn.miscTooling] = job.miscToolingCost;
            }
        }
    }
}
