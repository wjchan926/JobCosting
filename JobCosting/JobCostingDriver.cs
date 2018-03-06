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
            Dictionary<string, Job> jobList = new Dictionary<string, Job>();
            Job j1 = new Job("BCLF", "02328007-38");
            Job j2 = new Job("BCPN", "02343001-81");
            Job j3 = new Job("BDAB", "02240001-81");
            Job j4 = new Job("BDFJ", "01038014-38");
 
            jobList.Add(j1.salesOrder, j1);
            jobList.Add(j2.salesOrder, j2);
            jobList.Add(j3.salesOrder, j3);
            jobList.Add(j4.salesOrder, j4);


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
                foreach (KeyValuePair<string, Job> entry in jobList)
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

            // Write to excel, maybe with excel writer class
            // Test Print
            foreach (Job job in jobList.Values)
            {
                Console.WriteLine(job.customerName);
                Console.WriteLine(job.partNumber);
                Console.WriteLine(job.salesOrder);
                Console.WriteLine(job.salesRep);
                Console.WriteLine(job.productCost);
                Console.WriteLine(job.amountActualRevenue);
                Console.WriteLine(job.miscToolingCost);
                Console.WriteLine(job.freight);
                Console.WriteLine("\n");

            }

            Console.WriteLine("End of Main");
            
        }

    }
}
