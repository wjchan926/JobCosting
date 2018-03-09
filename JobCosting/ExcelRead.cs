using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace JobCosting
{
    /// <summary>
    /// Class that reads the Job Costing document
    /// </summary>
    sealed class ExcelRead
    {
        public Excel.Application myApp { get; set; } = null;
        public Excel.Workbook myBook { get; set; } = null;
        public Excel.Workbooks myBooks { get; set; } = null;
        public Excel.Worksheet mySheet { get; set; } = null;
        public Excel.Range myRange { get; private set; } = null;

        private string jobCostingPath = @"S:\JOB COSTING REFERENCE WORK SHEET.xlsm";
        
        /// <summary>
        /// Opens the Job Costing Document
        /// </summary>
        public void openDoc()
        {
            // Creates new instance of excel
            try
            {
                myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                myApp = new Excel.Application();          
            }
  
            // True to see new instance, false to hide
            myApp.Visible = true;
            myApp.DisplayAlerts = false;

            // Sets workbook to path specified
            // Try to open workbook.           
            try
            {
                myBooks = myApp.Workbooks;
                myBook = myBooks.Open(jobCostingPath);         
                mySheet = myBook.ActiveSheet;
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
        }

        /// <summary>
        /// Relinks the open excel document to this program
        /// </summary>
        public void reInitialize()
        {
            try
            {
                myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                myBooks = myApp.Workbooks;
                myBook = myApp.ActiveWorkbook;
                mySheet = myBook.ActiveSheet;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }         
        }

        /// <summary>
        /// Closes workbook and applicaiton.  Releases Objects
        /// Called with the Save and Close GUI Method
        /// </summary>
        public void close()
        {
            try
            {
                myBook.Close(true, Type.Missing, Type.Missing);
                myBooks.Close();
                myApp.Quit();
                myApp.DisplayAlerts = true;
                release();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }                 
        }

        /// <summary>
        /// Release objects
        /// </summary>
        public void release()
        {
            try
            {
                Marshal.ReleaseComObject(mySheet);
                Marshal.ReleaseComObject(myBooks);
                Marshal.ReleaseComObject(myBook);
                if (myRange != null) { Marshal.ReleaseComObject(myRange); }
                Marshal.ReleaseComObject(myApp);              
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Sets the range of of jobs to analyze with current selection
        /// </summary>
        public void setRange()
        {
            // Try to get selected range of cells from excel     
            try
            {
                myRange = myApp.Selection;
                Console.WriteLine("Range has been set");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Selection is not a valid range of cells");
                ConsoleWriter.WriteLine("Selection is not a valid range of cells");
            }
        }               
         
    }
}
