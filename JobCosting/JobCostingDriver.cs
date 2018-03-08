using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JobCosting
{
    static class JobCostingDriver
    {
        public static void Main()
        {
            Application.Run(new JobCostingGUI());
        }

        public static Dictionary<string,SuperJob> CostingDriver(ExcelRead jobCostingDoc)
        {
            if (jobCostingDoc == null)
            {
                jobCostingDoc.reInitialize();
            }

            Console.WriteLine("Main() Starting");
                                    
            // Create Conneciton Object
            QuickBooksConnector QBConnector = new QuickBooksConnector();

            // Connect
            QBConnector.connect();

            // Multithread objects
            Thread threadSO = new Thread(new ThreadStart(QuickBooksConnector.ThreadQuery.threadQuerySalesOrder));
            Thread threadCost = new Thread(new ThreadStart(QuickBooksConnector.ThreadQuery.threadCost));

            // Start threads
            try
            {
                threadSO.Start();
                threadCost.Start();
            }
            catch (ThreadStartException tse)
            {
                Console.WriteLine(tse.ToString());
            }

            // Wait for both threads to finish before continuing
            threadSO.Join();
            threadCost.Join();

            // For test puposes only
            // Create List to hold Job Objects
            // Key Should be SO#
            Dictionary<string, SuperJob> jobList = new Dictionary<string, SuperJob>();

            // Add All jobs highlighted in excel to dicitonary
            /**
              * 
              * Get all the sales orders and part numbers from Rows highlighted in excel using excel reader class
              * Create job objects with Sales Orders and PT#
              * Add to jobList
              * 
              * */
            Excel.Range myRange = jobCostingDoc.myRange;
            Excel.Worksheet exSheet = jobCostingDoc.mySheet;

            foreach (Excel.Range range in myRange.Rows)
            {                

                if(exSheet.Name != "StockingOrders")
                {                   
                    string soStr = (exSheet.Cells[range.Row, ExcelColumn.salesOrder]).Value.ToString().Substring(0,4);
                    string partNumberStr = (exSheet.Cells[range.Row, ExcelColumn.partNumber]).Value.ToString();
                    long orderQtyLong = (long)(exSheet.Cells[range.Row, ExcelColumn.orderQuantity]).Value;
                    
                    Job job = new Job(soStr, partNumberStr, orderQtyLong);

                    jobList.Add(job.salesOrder, job);
                }
                else
                {
                    string soStr = (exSheet.Cells[range.Row, ExcelColumn.salesOrder]).Value.ToString().Substring(0, 4);
                    string partNumberStr = (exSheet.Cells[range.Row, ExcelColumn.partNumber]).Value.ToString();
                    long orderQtyLong = (long)(exSheet.Cells[range.Row, ExcelColumn.orderQuantity]).Value;        

                    StockJob job = new StockJob(soStr, partNumberStr, orderQtyLong);
                    job.expectedRevenue = (double)(exSheet.Cells[range.Row, ExcelColumn.expectedAmount]).Value;
                    jobList.Add(job.salesOrder, job);
                }                   
            }
 
            // Map values from tables to job objects
            // Store all threads in list so we can check if closed.
            List<Thread> threadList = new List<Thread>();

            try
            {
                foreach (KeyValuePair<string, SuperJob> entry in jobList)
                {
                    // Find Row in Data Tables                    
                    string s = "SalesOrder='" + entry.Key + "'";
                    DataRow rowSO = QuickBooksConnector.result_SalesOrder.Select(s)[0];
                    DataRow rowItem = QuickBooksConnector.result_Cost.Rows.Find(entry.Value.partNumber);

                    // Map Values
                    entry.Value.customerName = rowSO["FullName"].ToString();
                    entry.Value.salesRep = rowSO["Rep"].ToString();

                    // This is the actual cost of the part
                    entry.Value.productCost = (decimal)rowItem["unit_cost_amt"];

                    // Run stored procedures for every job object in jobList
                    try
                    {
                        Thread thread = new Thread(new ParameterizedThreadStart(QuickBooksConnector.ThreadQuery.threadStoredProcedure));
                        threadList.Add(thread);
                        thread.Start(entry.Value);
                        Console.WriteLine("Starting Thread: " + entry.Key);
                    }
                    catch (ThreadStartException tse)
                    {
                        Console.WriteLine(tse.ToString());
                    }
                }
                Console.WriteLine("Properties mapped from SO and Item Tables");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            // Check if all threads have ended
            foreach (Thread thread in threadList)
            {
                thread.Join();
            }
            

            // Disconnect
            QBConnector.disconnect();

            // Finish calculations for jobs    
            foreach (KeyValuePair<string, SuperJob> entry in jobList)
            {  
                entry.Value.calculateFields();
            }

            Console.WriteLine("End of Main");

            return jobList;
          
            
        }

    }
}
