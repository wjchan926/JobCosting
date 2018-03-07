using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JobCosting
{
    sealed class ExcelRead
    {
        public Excel.Workbook myBook { get; set; } = null;
        public Excel.Application myApp { get; set; } = null;
        public Excel.Worksheet mySheet { get; set; } = null;
        public Excel.Range myRange { get; private set; } = null;

        private string jobCostingPath = @"S:\JOB COSTING REFERENCE WORK SHEET.xlsm";

        public void openDoc()
        {
            // Creates new instance of excel
            myApp = new Excel.Application();
  
            // True to see new instance, false to hide
            myApp.Visible = true;

            // Hide alerts
            myApp.DisplayAlerts = false;

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
            mySheet = myBook.ActiveSheet;
        }

        public void reInitialize()
        {            
            myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            myApp.DisplayAlerts = false;
            myBook = myApp.ActiveWorkbook;
            mySheet = myBook.ActiveSheet;           
        }

        /// <summary>
        /// Closes workbook and applicaiton.  Releases Objects
        /// Called with the Save and Close GUI Method
        /// </summary>
        public void close()
        {             
            try
            {                
                myBook.Close(true, null, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            myApp.DisplayAlerts = true;
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
                myRange = myApp.Selection;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Selection is not a valid range of cells");
            }
        }               
         
    }
}
