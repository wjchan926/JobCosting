using System;
using System.Data.Odbc;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Collections;
using System.Collections.Generic;

namespace JobCosting
{
    /// <summary>
    /// Connects QuickBooks to this application
    /// </summary>
    class QuickBooksConnector
    {
        public static OdbcConnection con { get; private set; }
        public static OdbcConnection conDSNLess { get; private set; }
        public static DataTable result_Cost { get; private set; }
        public static DataTable result_SalesOrder { get; private set; }

        /// <summary>
        /// Default Constructor
        /// </summary>
        public QuickBooksConnector()
        {
            result_Cost = new DataTable();
            result_SalesOrder = new DataTable();
        }

        /// <summary>
        /// Connect OdbcConnection to QuickBooks
        /// </summary>
        public void connect()
        {
            // Try to create QODBC Connection
            try
            {
                con = new OdbcConnection("Dsn=QuickBooks Data");

                con.Open(); // Open Connection, Required QB to be open
                Console.WriteLine("Connected to QB Thru QODBC");
                ConsoleWriter.WriteLine("Connected to QB Thru QODBC");
            }
            catch (Exception dbConnectionEx)
            {
                if (con != null)
                {
                    con.Dispose();
                }
                ConsoleWriter.WriteLine("Unable to Create Connection.");
                Console.WriteLine(dbConnectionEx.Message);
                throw;
            }

            // Try to Create DNSLess ODBC Connection   

            try
            {
                // Basically QB DSN Config file
                string fileDSN = @"\\MSW-FP1\Quickbooks\Company File 7_24_20\Marlin Steel Wire Products, LLC.QBW.DSN";

                // Extract line from file
                string[] lines = System.IO.File.ReadAllLines(fileDSN);

                string driver = "";
                string serverName = "";
                string commLinks = "";
                string databaseName = "";

                // Dynamically extract config infromation from FileDSN
                foreach (string line in lines)
                {
                    if (line.StartsWith("DRIVER="))
                    {
                        driver = line.Substring("DRIVER=".Length);
                        continue;
                    }
                    if (line.StartsWith("ServerName="))
                    {
                        serverName = line.Substring("ServerName=".Length); ;
                        continue;
                    }
                    if (line.StartsWith("CommLinks="))
                    {
                        commLinks = line.Substring("CommLinks=".Length); ;
                        continue;
                    }
                    if (line.StartsWith("DatabaseName="))
                    {
                        databaseName = line.Substring("DatabaseName=".Length); ;
                        continue;
                    }
                }

                conDSNLess = new OdbcConnection("ODBC; Driver={" + driver + "}; " +
                 "UID=JobCosting; " +
                 "PWD=M@rl1n; " +
                 "DatabaseName = " + databaseName + "; " +
                 "ServerName=" + serverName + "; " +
                 "AutoStop=NO; Integrated = NO; " +
                 "FILEDSN=" + fileDSN + ";" +
                 "Debug=NO; DisableMultiRowFetch=NO; CommLinks='" + commLinks + "'");

                conDSNLess.Open(); // Open Connection, Required QB to be open
                Console.WriteLine("Connected to QB Thru DSNLess Connection");
                ConsoleWriter.WriteLine("Connected to QB Thru DSNLess Connection");

            }
            catch (Exception dbConnectionEx)
            {
                if (conDSNLess != null)
                {
                    conDSNLess.Dispose();
                }

                Console.WriteLine(dbConnectionEx.Message);
                throw;
            }
        }

        /// <summary>
        /// Disconnects OdbcConnection from QuickBooks
        /// </summary>
        public void disconnect()
        {
            con.Close(); // Close Connection
            conDSNLess.Close(); // Close Connection
            Console.WriteLine("Disconnected from QB");
            ConsoleWriter.WriteLine("Disconnected from QB");
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

            public static void threadCost()
            {
                Console.WriteLine("Cost Query Started");
                // Create SQL statement for grabbing table data
                // Gets the Actual Material Cost from QuickBooks
                OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                    "SELECT name, unit_cost_amt, is_hidden " +
                    "FROM QBReportAdminGroup.v_lst_item ",
                    conDSNLess);

                // Store query results into DataTable Object
                dAdapter.Fill(result_Cost);

                // Set Primary Key
                DataColumn[] key = new DataColumn[1];
                key[0] = result_Cost.Columns["name"];
                result_Cost.PrimaryKey = key;

                Console.WriteLine("Cost Table Filled");
                ConsoleWriter.WriteLine("Cost Table Filled");
            }

            /// <summary>
            /// Fills result_SalesOrder with the ODBC DSNLess Connections Tables from QB
            /// </summary>
            public static void threadQuerySalesOrder()
            {

                Console.WriteLine("SO Table Query Started");
                // Create SQL statement for grabbing table data
                OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                    "SELECT QBReportAdminGroup.v_lst_customer_fullname.name as 'SalesOrder', QBReportAdminGroup.v_lst_customer_fullname.full_name as 'FullName', QBReportAdminGroup.v_lst_sales_rep.initials as 'Rep'" +
                    "FROM QBReportAdminGroup.v_lst_customer INNER JOIN QBReportAdminGroup.v_lst_customer_fullname ON QBReportAdminGroup.v_lst_customer.name = QBReportAdminGroup.v_lst_customer_fullname.name " +
                    "INNER JOIN QBReportAdminGroup.v_lst_sales_rep ON QBReportAdminGroup.v_lst_customer.sales_rep_id = QBReportAdminGroup.v_lst_sales_rep.id",
                    conDSNLess);

                // Store query results into DataTable Object
                dAdapter.Fill(result_SalesOrder);

                Console.WriteLine("SO Table Filled");
                ConsoleWriter.WriteLine("SO Table Filled");
            }

            public static void tableWriter(DataTable dataTable)
            {
                DataTableReader dataTableReader = new DataTableReader(dataTable);
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
            /// Public overloaded method that accepts an object ofr ParameterizedThreadStart
            /// </summary>
            /// <param name="joblist"></param>
            public static void threadStoredProcedure(object job)
            {
                threadStoredProcedure((SuperJob)job);
            }

            /// <summary>
            /// Runs the stored procedure Report in QB for the job.  This needs ODBC conneciton, unfortunately.
            /// </summary>
            /// <param name="jobList"></param>
            private static void threadStoredProcedure(SuperJob job)
            {
                Console.WriteLine("Stored Procedure Started");
                ConsoleWriter.WriteLine("Stored Procedure(s) Started");

                // Create SQL statement for grabbing table data
                DataTable result_StoredProcedure = new DataTable();

                try // Try to query
                {
                    string customer = job.customerName;
                    OdbcDataAdapter dAdapter = new OdbcDataAdapter(
                      "sp_report JobProfitabilityDetail " +
                      "show RowData, AmountActualCost, AmountActualRevenue, AmountDifferenceActual " +
                      "parameters DateMacro = 'All', EntityFilterFullNames = '" + customer + "'",
                      con);

                    // Store query results into DataTable Object
                    dAdapter.Fill(result_StoredProcedure);

                    // Set Primary Key
                    // Replace Null Values in RowData Column
                    // Stored Procedure $ are of type Decimal
                    foreach (DataRow row in result_StoredProcedure.Rows)
                    {
                        if (row["RowData"] is System.DBNull)
                        {
                            row["RowData"] = "NO DATA" + result_StoredProcedure.Rows.IndexOf(row);
                        }
                    }

                    DataColumn[] key = new DataColumn[1];
                    key[0] = result_StoredProcedure.Columns["RowData"];
                    result_StoredProcedure.PrimaryKey = key;

                    // Map data to job objects
                    mapData(job, result_StoredProcedure);
                }

                catch (OdbcException objEx)
                {
                    Console.WriteLine(objEx.Message);
                }
            }

            /// <summary>
            /// Maps data to job entries
            /// </summary>
            /// <param name="job"></param> Job analyzed
            /// <param name="result_StoredProcedure"></param> Table that has all stored procedure data
            private static void mapData(SuperJob job, DataTable result_StoredProcedure)
            {
                try // Try to map bad cost data
                {
                    job.badCostData = (decimal)result_StoredProcedure.Rows.Find(job.partNumber)["AmountActualCost_1"];
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    Console.WriteLine("No material data found for: " + job.customerName);
                    ConsoleWriter.WriteLine("No material data found for: " + job.customerName);
                }

                // Map total Costs
                job.amountActualCost = (decimal)result_StoredProcedure.Rows[result_StoredProcedure.Rows.Count - 1]["AmountActualCost_1"];

                // Map total Revenue
                job.amountActualRevenue = (decimal)result_StoredProcedure.Rows[result_StoredProcedure.Rows.Count - 1]["AmountActualRevenue_1"];

                try // Try to Map freight if found
                {
                    job.freight = (decimal)result_StoredProcedure.Rows.Find("Freight")["AmountActualRevenue_1"];
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    Console.WriteLine("No freight data found for: " + job.customerName);
                    ConsoleWriter.WriteLine("No freight data found for: " + job.customerName);
                }

                try // Try to map msc Tooling if found
                {
                    job.miscToolingCost = (decimal)result_StoredProcedure.Rows.Find("MISC TOOLING")["AmountActualRevenue_1"];
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    Console.WriteLine("No tooling data found for: " + job.customerName);
                    ConsoleWriter.WriteLine("No tooling data found for: " + job.customerName);
                }
            }
        }


    }


}
