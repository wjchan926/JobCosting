using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace JobCosting
{
    /// <summary>
    /// Wrties the data and colors in Excel Job Costing Document
    /// </summary>
    static class ExcelWrite
    {       
        // Write Methods
        // May need to change
        /// <summary>
        /// Writes job data to the excel sheet
        /// </summary>
        /// <param name="job"></param> Job to be analyzed
        public static void writeJobData(Dictionary<string, SuperJob> jobList, Excel.Range selectedRange, Excel.Worksheet mySheet)
        {
            foreach (Excel.Range range in selectedRange.Rows)
            {
                dynamic soValue = mySheet.Cells[range.Row, ExcelColumn.salesOrder].Value;
                string soStr = soValue.ToString().Substring(0, 4);

                mySheet.Cells[range.Row, ExcelColumn.salesRep] = jobList[soStr].salesRep;
                mySheet.Cells[range.Row, ExcelColumn.actualCost] = string.Format("{0:C}", jobList[soStr].amountActualCost);
                mySheet.Cells[range.Row, ExcelColumn.actualRevenue] = string.Format("{0:C}", jobList[soStr].amountActualRevenue);
                mySheet.Cells[range.Row, ExcelColumn.difference] = string.Format("{0:C}", jobList[soStr].difference);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin] = string.Format("{0:0.00%}",jobList[soStr].grossMargin);
                mySheet.Cells[range.Row, ExcelColumn.unitHigh] = string.Format("{0:C}", jobList[soStr].unitHigh);
                mySheet.Cells[range.Row, ExcelColumn.unitMed] = string.Format("{0:C}", jobList[soStr].unitMed);
                mySheet.Cells[range.Row, ExcelColumn.unitLow] = string.Format("{0:C}", jobList[soStr].unitLow);
                mySheet.Cells[range.Row, ExcelColumn.unitFloor] = string.Format("{0:C}", jobList[soStr].unitFloor);
                mySheet.Cells[range.Row, ExcelColumn.freight] = string.Format("{0:C}", jobList[soStr].freight);
                mySheet.Cells[range.Row, ExcelColumn.marlinFreight] = string.Format("{0:C}", jobList[soStr].marlinFreight);
                mySheet.Cells[range.Row, ExcelColumn.miscTooling] = string.Format("{0:C}", jobList[soStr].miscToolingCost);
                mySheet.Cells[range.Row, ExcelColumn.costToCure] = string.Format("{0:C}", jobList[soStr].costToCure);
                mySheet.Cells[range.Row, ExcelColumn.healingFactor] = string.Format("{0:C}", jobList[soStr].healingFactor);

                formatJobDoc(jobList[soStr], range, mySheet);
                ConsoleWriter.WriteLine(jobList[soStr].partNumber + " | " + jobList[soStr].customerName + " Data Written to Excel.");
            }
        }

        /// <summary>
        /// Colors the cells based on job costing values
        /// </summary>
        /// <param name="job"></param> Job analyzed
        /// <param name="range"></param> row that the job is in
        /// <param name="mySheet"></param> current worksheet
        private static void formatJobDoc(SuperJob job, Excel.Range range, Excel.Worksheet mySheet)
        {
            if (job.salesRep == "Not Fully Invoiced" || job.salesRep == "No Revenue for Job" || job.salesRep == "TimeClock Not Imported")
            {
                mySheet.Cells[range.Row, ExcelColumn.salesRep].Interior.Color = Color.FromArgb(96, 96, 96);
            }
            else
            {
                mySheet.Cells[range.Row, ExcelColumn.salesRep].Interior.ColorIndex = Excel.Constants.xlNone;
            }

            mySheet.Cells[range.Row, ExcelColumn.actualCost].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.actualRevenue].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.difference].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.freight].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.marlinFreight].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.miscTooling].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.costToCure].Interior.ColorIndex = Excel.Constants.xlNone;
            mySheet.Cells[range.Row, ExcelColumn.healingFactor].Interior.ColorIndex = Excel.Constants.xlNone;

            if (job.salesRep == "Not Fully Invoiced" || job.salesRep == "No Revenue for Job" || job.salesRep == "TimeClock Not Imported")
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(96, 96, 96);
            }
            else if(job.grossMargin <= -.25)
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(255, 182, 193);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Font.Color = Color.FromArgb(178, 34, 34);
            }
            else if (job.grossMargin <= 0)
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(255, 204, 229);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Font.Color = Color.FromArgb(225, 0, 127);
            }
            else if (job.grossMargin <.25)
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(255, 128, 0);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Font.Color = Color.FromArgb(0, 0, 0);
            }
            else if (job.grossMargin < .42)
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(255, 255, 0);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Font.Color = Color.FromArgb(0, 0, 0);
            }
            else if (job.grossMargin >= .42)
            {
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Interior.Color = Color.FromArgb(182, 255, 193);
                mySheet.Cells[range.Row, ExcelColumn.grossMargin].Font.Color = Color.FromArgb(20, 100, 20);
            }

            if (job.salesRep == "Not Fully Invoiced" || job.salesRep == "No Revenue for Job" || job.salesRep == "TimeClock Not Imported")
            {
                mySheet.Cells[range.Row, ExcelColumn.unitHigh].Interior.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitMed].Interior.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitLow].Interior.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitFloor].Interior.Color = Color.FromArgb(96, 96, 96);

                mySheet.Cells[range.Row, ExcelColumn.unitHigh].Font.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitMed].Font.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitLow].Font.Color = Color.FromArgb(96, 96, 96);
                mySheet.Cells[range.Row, ExcelColumn.unitFloor].Font.Color = Color.FromArgb(96, 96, 96);
            }
            else
            {
                mySheet.Cells[range.Row, ExcelColumn.unitHigh].Interior.ColorIndex = Excel.Constants.xlNone;
                mySheet.Cells[range.Row, ExcelColumn.unitMed].Interior.ColorIndex = Excel.Constants.xlNone;
                mySheet.Cells[range.Row, ExcelColumn.unitLow].Interior.ColorIndex = Excel.Constants.xlNone;
                mySheet.Cells[range.Row, ExcelColumn.unitFloor].Interior.ColorIndex = Excel.Constants.xlNone;

                mySheet.Cells[range.Row, ExcelColumn.unitHigh].Font.Color = Color.FromArgb(0, 0, 0);
                mySheet.Cells[range.Row, ExcelColumn.unitMed].Font.Color = Color.FromArgb(0, 0, 0);
                mySheet.Cells[range.Row, ExcelColumn.unitLow].Font.Color = Color.FromArgb(0, 0, 0);
                mySheet.Cells[range.Row, ExcelColumn.unitFloor].Font.Color = Color.FromArgb(0, 0, 0);
            }

        }
    }
}
