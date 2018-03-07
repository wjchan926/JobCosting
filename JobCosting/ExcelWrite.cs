using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JobCosting
{
    static class ExcelWrite
    {       
        public static Excel.Application myApp { get; private set; } = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        public static Excel.Workbook myBook { get; private set; } = myApp.ActiveWorkbook;
        public static Excel.Worksheet mySheet { get; private set; } = myBook.ActiveSheet;

        // Write Methods
        // May need to change
        /// <summary>
        /// Writes job data to the excel sheet
        /// </summary>
        /// <param name="job"></param> Job to be analyzed
        public static void writeJobData(Dictionary<string, SuperJob> jobList, Excel.Range selectedRange)
        {
            myApp  = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            myBook  = myApp.ActiveWorkbook;
            mySheet = myBook.ActiveSheet;

            foreach (Excel.Range range in selectedRange.Rows)
            {
                string soStr = (mySheet.Cells[range.Row, ExcelColumn.salesOrder]).Value.ToString().Substring(0, 4);

                mySheet.Cells[range.Row, ExcelColumn.salesRep] = jobList[soStr].salesRep;
                mySheet.Cells[range.Row, ExcelColumn.actualCost] = jobList[soStr].amountActualCost;
                mySheet.Cells[range.Row, ExcelColumn.actualRevenue] = jobList[soStr].amountActualRevenue;
                mySheet.Cells[range.Row, ExcelColumn.difference] = jobList[soStr].difference;
                mySheet.Cells[range.Row, ExcelColumn.grossMargin] = jobList[soStr].grossMargin;
                mySheet.Cells[range.Row, ExcelColumn.unitHigh] = jobList[soStr].unitHigh;
                mySheet.Cells[range.Row, ExcelColumn.unitMed] = jobList[soStr].unitMed;
                mySheet.Cells[range.Row, ExcelColumn.unitLow] = jobList[soStr].unitLow;
                mySheet.Cells[range.Row, ExcelColumn.unitFloor] = jobList[soStr].unitFloor;
                mySheet.Cells[range.Row, ExcelColumn.freight] = jobList[soStr].freight;
                mySheet.Cells[range.Row, ExcelColumn.marlinFreight] = jobList[soStr].marlinFreight;
                mySheet.Cells[range.Row, ExcelColumn.miscTooling] = jobList[soStr].miscToolingCost;
            }
        }
    }
}
