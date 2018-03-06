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
        ExcelReadWrite jobCostingDoc;

        public JobCostingGUI()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            jobCostingDoc.release();
        }

        private void openbtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc = new ExcelReadWrite();
        }

        private void analyzebtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc.setRange();
            JobCostingDriver.CostingDriver(jobCostingDoc);           
        }

        private void closebtn_Click(object sender, EventArgs e)
        {
            jobCostingDoc.close();
        }
    }
}
