using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JobCosting
{
    public partial class JobCostingGUI : Form
    {
        ExcelRead jobCostingDoc;

        public JobCostingGUI()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void openbtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc = new ExcelRead();
            jobCostingDoc.openDoc();
        }

        private void analyzebtn_Click(object sender, EventArgs e)
        {
            if (jobCostingDoc.myApp == null)
            {
                jobCostingDoc = new ExcelRead();        
            }

            jobCostingDoc.reInitialize();
            jobCostingDoc.setRange();
            Dictionary<string, SuperJob> jobList = JobCostingDriver.CostingDriver(jobCostingDoc);
            ExcelWrite.writeJobData(jobList, jobCostingDoc.myRange); 
        }

        private void closebtn_Click(object sender, EventArgs e)
        {
            if (jobCostingDoc.myApp == null)
            {
                jobCostingDoc = new ExcelRead();

            }
            jobCostingDoc.reInitialize();

            jobCostingDoc.close();
        }

        private void exitbtn_Click(object sender, EventArgs e)
        {
            if (jobCostingDoc != null) { jobCostingDoc.release(); }
            Application.Exit();
        }
    }
}
