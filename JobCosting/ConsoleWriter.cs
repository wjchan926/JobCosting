using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JobCosting
{
    static class ConsoleWriter
    {
        private static Control textbox;
        private static JobCostingGUI jobCostingGUI;

        public static void setTextBox(Control outputTb)
        {
            textbox = outputTb;
        }

        public static void setGUI(JobCostingGUI GUI)
        {
            jobCostingGUI = GUI;
        }

        public static void WriteLine(string value)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(textbox.Text);
            sb.Append(value);

            if (jobCostingGUI.InvokeRequired)
            {
                jobCostingGUI.Invoke(new Action<string>(WriteLine), new object[] { value });
                return;
            } 

            textbox.Text = sb.ToString() + Environment.NewLine;  
        }
        
        public static void AppendTextBox(string value)
        {
            if (jobCostingGUI.InvokeRequired)
            {
                jobCostingGUI.Invoke(new Action<string>(AppendTextBox), new object[] { value });
                return;
            }
            textbox.Text += value;
        }

    }
}
