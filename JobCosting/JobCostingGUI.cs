using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Deployment.Application;

namespace JobCosting
{
    public partial class JobCostingGUI : Form
    {
        ExcelRead jobCostingDoc = new ExcelRead();
        string version;

        public JobCostingGUI()
        {
            InitializeComponent();
            try
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();                
            }
            catch (Exception)
            {
                
            }
            if (version == null)
            {
                this.Text = "Job Costing Tool VDebug";
            }
            else
            {
                this.Text = "Job Costing Tool V" + version;
            }       


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            TopMost = true;

            // Set output to textbox
            ConsoleWriter.setGUI(this);
            ConsoleWriter.setTextBox(outputTb);
            ConsoleWriter.WriteLine("Please Ensure QuickBooks Application is Open.");
            ConsoleWriter.WriteLine("Job Costing Analyzer Tool Starting.");
        }


        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ConsoleWriter.WriteLine("Analyzing Jobs.");

            jobCostingDoc.reInitialize();
            jobCostingDoc.setRange();

            if (jobCostingDoc.myRange != null)
            {
                Dictionary<string, SuperJob> jobList = JobCostingDriver.CostingDriver(jobCostingDoc);

                Excel.Range myRange = jobCostingDoc.myRange;
                Excel.Worksheet mySheet = jobCostingDoc.mySheet;

                ExcelWrite.writeJobData(jobList, myRange, mySheet);

                Marshal.ReleaseComObject(myRange);
                Marshal.ReleaseComObject(mySheet);
            }

            ConsoleWriter.WriteLine("Analysis of Jobs Complete.");
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //Do something when the process finishes
           // MessageBox.Show("Done Counting!");
        }

        //This fires on the UI Thread so you can update the textbox here
        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            outputTb.Update();
            ////now we get some data in the e.UserState object
            ////this can be any object, in this example it is just a string
            ////that is the product of the loop count times two
            //textBox1.AppendText(Environment.NewLine + e.UserState.ToString());
        }


        /// <summary>
        /// Opens Job Costing Document Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openbtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc.openDoc();            
            ConsoleWriter.WriteLine("Job Costing Document Opened");            
        }

        /// <summary>
        /// Analyze Jobs Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void analyzebtn_Click(object sender, EventArgs e)
        {
            backgroundWorker.RunWorkerAsync();  
        }

        /// <summary>
        /// Save  and close Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closebtn_Click(object sender, EventArgs e)
        {
            
            ConsoleWriter.WriteLine("Saving and Closing Job Costing Doc.");            

            try
            {
                // Open, but didnt initialize
                jobCostingDoc.reInitialize();
                jobCostingDoc.close();
                ConsoleWriter.WriteLine("Job Costing Document Closed.");
            }
            catch (Exception ex)
            {
                // No Document
                // Do nothing
                Console.WriteLine(ex.Message);
            }            
        }

        /// <summary>
        /// Exit Application Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitbtn_Click(object sender, EventArgs e)
        {
            if (jobCostingDoc != null) { jobCostingDoc.release(); }
            Application.Exit();
        }

        private void outputTb_TextChanged(object sender, EventArgs e)
        {
            outputTb.SelectionStart = outputTb.Text.Length;
            outputTb.ScrollToCaret();
            outputTb.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            outputTb.Clear();
        }
    }
}
