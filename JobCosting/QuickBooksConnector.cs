using System;
using System.Data.Odbc;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Collections;
using System.Collections.Generic;

namespace JobCosting
{
    class QuickBooksConnector
    {
        public static OdbcConnection con { get; private set; }
        public static DataTable result_SalesOrder { get; private set; }
        public static DataTable result_ItemInventoryAssembly { get; private set; }
      //  public static DataTable result_StoredProcedure { get; private set; }

        public QuickBooksConnector()
        {
            result_SalesOrder = new DataTable();
            result_ItemInventoryAssembly = new DataTable();            
        }

        /// <summary>
        /// Conenct OdbcConnection to QuickBooks
        /// </summary>
        public void connect()
        {
            try
            {
                con = new OdbcConnection("Dsn=QuickBooks Data");
                con.Open(); // Open Connection, Required QB to be open
                Console.WriteLine("Connected to QB");
            }
            catch (Exception dbConnectionEx)
            {
                if (con != null)
                {
                    con.Dispose();
                }

                Console.WriteLine(dbConnectionEx.Message);
                throw;
            }

            //      DataTableReader reader = new DataTableReader(result_Item);

            //        tableWriter(reader);

        }

        /// <summary>
        /// Disconnects OdbcConnection from QuickBooks
        /// </summary>
        public void disconnect()
        {
            con.Close(); // Close Connection
            Console.WriteLine("Disconnected from QB");
        }

        /// <summary>
        /// Prints the table contents to the console for debugging purposes
        /// </summary>
        /// <param name="dataTableReader"></param>
        /// <param name="readColumns"></param>
        public void tableWriter(DataTableReader dataTableReader)
        {
            // Read table, while there is still a record
            while (dataTableReader.Read())
            {
                try
                {
                    Console.WriteLine(dataTableReader.GetString(0));
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine(e.Message);
                }
            }
        }


        /// <summary>
        /// Nested class for hyperthreading the queries
        /// </summary>
        public static class ThreadQuery
        {            
            /// <summary>
            /// Fills result_SalesOrder with the SalesOrder Table from QB
            /// </summary>
            public static void threadQuerySalesOrder()
            {

                Console.WriteLine("SO Table Query Started");
                // Create SQL statement for grabbing table data
                OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                    "SELECT CustomerRefFullName, SalesRepRefFullName, RefNumber, IsFullyInvoiced, IsManuallyClosed " +
                    "FROM SalesOrder ",
                    con);

                // Store query results into DataTable Object
                dAdapter.Fill(result_SalesOrder);

                // Set Primary Key
                DataColumn[] key = new DataColumn[1];
                key[0] = result_SalesOrder.Columns["RefNumber"];
                result_SalesOrder.PrimaryKey = key;

                Console.WriteLine("SO Table Filled");
         
            }
            
            /// <summary>
            /// Fills result_ItemInventoryAssembly with ItemInventoryAssembly Table from QB
            /// </summary>
            public static void threadQueryItemInventoryAssembly()
            {
                Console.WriteLine("Item Table Query Started");

                // Create SQL statement for grabbing table data         
                OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                    "SELECT FullName, AverageCost, IsActive " +
                    "FROM ItemInventoryAssembly " +
                    "WHERE IsActive=True", 
                    con);

                // Store query results into DataTable Object
                dAdapter.Fill(result_ItemInventoryAssembly);

                // Set Primary Key
                DataColumn[] key = new DataColumn[1];
                key[0] = result_ItemInventoryAssembly.Columns["FullName"];
                result_ItemInventoryAssembly.PrimaryKey = key;

                Console.WriteLine("Item Table Filled");   
            }


            /// <summary>
            /// Public overloaded method that accepts an object ofr ParameterizedThreadStart
            /// </summary>
            /// <param name="joblist"></param>
            public static void threadStoredProcedure(object job)
            {    
                threadStoredProcedure((Job)job);
            }

            /// <summary>
            /// Runs the stored procedure Report in QB for the job
            /// </summary>
            /// <param name="jobList"></param>
            private static void threadStoredProcedure(Job job)
            {
                Console.WriteLine("Stored Procedure Started");
                // Create SQL statement for grabbing table data
                DataTable result_StoredProcedure = new DataTable();

                try
                {
                    string customer = job.customerName;
                    OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                      "sp_report JobProfitabilityDetail " +
                      "show Label, AmountActualCost, AmountActualRevenue, AmountDifferenceActual " +
                      "parameters DateMacro = 'All', EntityFilterFullNames = '"+ customer + "'",
                      con);                    

                    // Store query results into DataTable Object
                    dAdapter.Fill(result_StoredProcedure);

                    // Set Primary Key
                    // Replace Null Values in Label Column
                    // Stored Procedure $ are of type Decimal
                    foreach(DataRow row in result_StoredProcedure.Rows)
                    {
                        if (row["Label"] is System.DBNull)
                        {
                            row["Label"] = "NO DATA" + result_StoredProcedure.Rows.IndexOf(row);
                        }
                    }

                    DataColumn[] key = new DataColumn[1];
                    key[0] = result_StoredProcedure.Columns["Label"]; 
                    result_StoredProcedure.PrimaryKey = key;

                    // Map data to job objects
                    job.amountActualCost = (decimal)result_StoredProcedure.Rows.Find("TOTAL")["AmountActualCost_1"];
                    job.amountActualRevenue = (decimal)result_StoredProcedure.Rows.Find("TOTAL")["AmountActualRevenue_1"];
                    try
                    {
                        job.freight = (decimal)result_StoredProcedure.Rows.Find("Freight (Freight Out)")["AmountActualRevenue_1"];
                    }
                    catch (Exception)
                    {
                        
                    }
                    job.marlinFreight = job.freight / (decimal)1.75;
                    //job.miscToolingCost = (decimal)result_StoredProcedure.Rows.Find("MISC TOOLING(One - time charge - engineering time, check fixtures, tooling, qualit...")["AmountActualRevenue_1"];

                    //Console.WriteLine(job.customerName);
                    //Console.WriteLine(job.averageCost);
                    //Console.WriteLine(job.amountActualCost);
                    //Console.WriteLine(job.amountActualRevenue);
                      
                }
                catch (OdbcException objEx)
                {
                    Console.WriteLine(objEx.Message);
                }                
            }
        }

        
    }


}
