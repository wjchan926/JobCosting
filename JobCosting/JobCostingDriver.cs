using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Data;

namespace JobCosting
{
    class JobCostingDriver
    {
        public static void Main()
        {
            Console.WriteLine("Main() Starting");
                                    
            // Create Conneciton Object
            QuickBooksConnector QBConnector = new QuickBooksConnector();

            // Connect
            QBConnector.connect();
      
            // Multithread objects
            Thread t1 = new Thread(new ThreadStart(QuickBooksConnector.ThreadQuery.threadQuerySalesOrder));
            Thread t2 = new Thread(new ThreadStart(QuickBooksConnector.ThreadQuery.threadQueryItemInventoryAssembly));


            // Start threads
            try
            {
                t1.Start();
                t2.Start();
            }
            catch (ThreadStartException tse)
            {
                Console.WriteLine(tse.ToString());
            }

            // Wait for both threads to finish before continuing
            t1.Join();
            t2.Join();

            // Create List to hold Job Objects
            // Key Should be SO#
            Dictionary<int,Job> jobList = new Dictionary<int, Job>();
            Job j1 = new Job(20633, "02328007-38");
            Job j2 = new Job(20745, "02343001-81");
            Job j3 = new Job(21023, "02240001-81");
            Job j4 = new Job(21160, "01038014-38");
            Job j5 = new Job(21150, "01698008-31");
            Job j6 = new Job(21025, "00531003-34");
            Job j7 = new Job(21260, "01145001-38");
            Job j8 = new Job(21204, "02400001-38");

            jobList.Add(j1.salesOrder, j1);
            jobList.Add(j2.salesOrder, j2);
            jobList.Add(j3.salesOrder, j3);
            jobList.Add(j4.salesOrder, j4);
            jobList.Add(j5.salesOrder, j5);
            jobList.Add(j6.salesOrder, j6);
            jobList.Add(j7.salesOrder, j7);
            jobList.Add(j8.salesOrder, j8);
     
            // Add All jobs highlighted in excel to dicitonary
            /**
             * 
             * Get all the sales orders and part numbers from Rows highlighted in excel using excel reader class
             * Create job objects with Sales Orders and PT#
             * Add to jobList
             * 
             * */

            // Map values from tables to job objects
            // Store all threads in list so we can check if closed.
            List<Thread> threadList = new List<Thread>();

            try
            {
                foreach (KeyValuePair<int, Job> entry in jobList)
                {
                    // Find Row in Data Tables
                    DataRow rowSO = QuickBooksConnector.result_SalesOrder.Rows.Find(entry.Key);
                    DataRow rowItem = QuickBooksConnector.result_ItemInventoryAssembly.Rows.Find(entry.Value.partNumber);

                    // Map Values
                    
                    entry.Value.customerName = rowSO["CustomerRefFullName"].ToString();
                    entry.Value.salesRep = rowSO["SalesRepRefFullName"].ToString();
                    entry.Value.isFullyInvoiced = (bool)rowSO["IsFullyInvoiced"];
                    entry.Value.isManuallyClosed = (bool)rowSO["IsManuallyClosed"];

                    // This will be overwritten by the stored provedure if not in StockingOrders Sheet
                    entry.Value.averageCost = (decimal)rowItem["AverageCost"];

                    // Run stored procedures for every job object in jobList
                    try
                    {
                        Thread thread = new Thread(new ParameterizedThreadStart(QuickBooksConnector.ThreadQuery.threadStoredProcedure));
                        threadList.Add(thread);
                        thread.Start(entry.Value);
                    }
                    catch (ThreadStartException tse)
                    {
                        Console.WriteLine(tse.ToString());
                    }
                }
                Console.WriteLine("Properties mapped from SO and Item Tables, threading started");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            
            // Check if all threads have ended
            foreach (Thread thread in threadList)
            {
                thread.Join();
            }
            // Disconnect
            QBConnector.disconnect();
            
            // Write to excel, maybe with excel writer class
            //Test Print
            foreach(Job job in jobList.Values)
            {
                Console.WriteLine(job.customerName);
                Console.WriteLine(job.partNumber);
                Console.WriteLine(job.salesOrder);
                Console.WriteLine(job.salesRep);
                Console.WriteLine(job.averageCost);
                Console.WriteLine(job.amountActualCost);
                Console.WriteLine(job.amountActualRevenue);
                Console.WriteLine(job.isFullyInvoiced);
                Console.WriteLine(job.isManuallyClosed);
                Console.WriteLine("\n");

            }

            Console.WriteLine("End of Main");
            
        }

    }
}
