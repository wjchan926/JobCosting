namespace JobCosting
{
    partial class JobCostingGUI
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openbtn = new System.Windows.Forms.Button();
            this.analyzebtn = new System.Windows.Forms.Button();
            this.closebtn = new System.Windows.Forms.Button();
            this.exitbtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.outputTb = new System.Windows.Forms.TextBox();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openbtn
            // 
            this.openbtn.Location = new System.Drawing.Point(12, 12);
            this.openbtn.Name = "openbtn";
            this.openbtn.Size = new System.Drawing.Size(175, 52);
            this.openbtn.TabIndex = 0;
            this.openbtn.Text = "Open Job Costing Document";
            this.openbtn.UseVisualStyleBackColor = true;
            this.openbtn.Click += new System.EventHandler(this.openbtn_Click);
            // 
            // analyzebtn
            // 
            this.analyzebtn.Location = new System.Drawing.Point(12, 70);
            this.analyzebtn.Name = "analyzebtn";
            this.analyzebtn.Size = new System.Drawing.Size(175, 52);
            this.analyzebtn.TabIndex = 1;
            this.analyzebtn.Text = "Analyze Selected Jobs";
            this.analyzebtn.UseVisualStyleBackColor = true;
            this.analyzebtn.Click += new System.EventHandler(this.analyzebtn_Click);
            // 
            // closebtn
            // 
            this.closebtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.closebtn.Location = new System.Drawing.Point(12, 128);
            this.closebtn.Name = "closebtn";
            this.closebtn.Size = new System.Drawing.Size(175, 23);
            this.closebtn.TabIndex = 2;
            this.closebtn.Text = "Save and Close Job Doc";
            this.closebtn.UseVisualStyleBackColor = true;
            this.closebtn.Click += new System.EventHandler(this.closebtn_Click);
            // 
            // exitbtn
            // 
            this.exitbtn.Location = new System.Drawing.Point(12, 157);
            this.exitbtn.Name = "exitbtn";
            this.exitbtn.Size = new System.Drawing.Size(175, 23);
            this.exitbtn.TabIndex = 4;
            this.exitbtn.Text = "Close Application Tool";
            this.exitbtn.UseVisualStyleBackColor = true;
            this.exitbtn.Click += new System.EventHandler(this.exitbtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.outputTb);
            this.groupBox1.Location = new System.Drawing.Point(193, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(365, 168);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Console";
            // 
            // outputTb
            // 
            this.outputTb.Location = new System.Drawing.Point(6, 19);
            this.outputTb.Multiline = true;
            this.outputTb.Name = "outputTb";
            this.outputTb.ReadOnly = true;
            this.outputTb.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.outputTb.Size = new System.Drawing.Size(353, 114);
            this.outputTb.TabIndex = 0;
            this.outputTb.TextChanged += new System.EventHandler(this.outputTb_TextChanged);
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(273, 139);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Clear Console";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // JobCostingGUI
            // 
            this.AcceptButton = this.analyzebtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 187);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.exitbtn);
            this.Controls.Add(this.closebtn);
            this.Controls.Add(this.analyzebtn);
            this.Controls.Add(this.openbtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "JobCostingGUI";
            this.Text = "Job Costing Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button openbtn;
        private System.Windows.Forms.Button analyzebtn;
        private System.Windows.Forms.Button closebtn;
        private System.Windows.Forms.Button exitbtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox outputTb;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.Button button1;
    }
}

