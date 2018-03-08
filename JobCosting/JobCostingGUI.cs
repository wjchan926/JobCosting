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

namespace JobCosting
{
    public partial class JobCostingGUI : Form
    {
        ExcelRead jobCostingDoc = new ExcelRead();

        public JobCostingGUI()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void openbtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc.openDoc();
        }

        private void analyzebtn_Click(object sender, EventArgs e)
        {
            
            jobCostingDoc.reInitialize();
            jobCostingDoc.setRange();

            if (jobCostingDoc.myRange != null)
            {
                Dictionary<string, SuperJob> jobList = JobCostingDriver.CostingDriver(jobCostingDoc);

                ExcelWrite.writeJobData(jobList, jobCostingDoc.myRange, jobCostingDoc.mySheet);
            }
        }

        private void closebtn_Click(object sender, EventArgs e)
        {
            try
            {
                // Open, but didnt initialize
                jobCostingDoc.reInitialize();
                jobCostingDoc.close();
            }
            catch (Exception ex)
            {
                // No Document
                // Do nothing
                Console.WriteLine(ex.Message);
            }            
        }

        private void exitbtn_Click(object sender, EventArgs e)
        {
            if (jobCostingDoc != null) { jobCostingDoc.release(); }
            Application.Exit();
        }
    }
}
