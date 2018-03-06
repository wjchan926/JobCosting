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
            this.cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openbtn
            // 
            this.openbtn.Location = new System.Drawing.Point(12, 12);
            this.openbtn.Name = "openbtn";
            this.openbtn.Size = new System.Drawing.Size(175, 23);
            this.openbtn.TabIndex = 0;
            this.openbtn.Text = "Open Job Costing Document";
            this.openbtn.UseVisualStyleBackColor = true;
            this.openbtn.Click += new System.EventHandler(this.openbtn_Click);
            // 
            // analyzebtn
            // 
            this.analyzebtn.Location = new System.Drawing.Point(12, 41);
            this.analyzebtn.Name = "analyzebtn";
            this.analyzebtn.Size = new System.Drawing.Size(175, 23);
            this.analyzebtn.TabIndex = 1;
            this.analyzebtn.Text = "Analyze Selected Jobs";
            this.analyzebtn.UseVisualStyleBackColor = true;
            this.analyzebtn.Click += new System.EventHandler(this.analyzebtn_Click);
            // 
            // closebtn
            // 
            this.closebtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.closebtn.Location = new System.Drawing.Point(12, 70);
            this.closebtn.Name = "closebtn";
            this.closebtn.Size = new System.Drawing.Size(174, 23);
            this.closebtn.TabIndex = 2;
            this.closebtn.Text = "Save and Close";
            this.closebtn.UseVisualStyleBackColor = true;
            this.closebtn.Click += new System.EventHandler(this.closebtn_Click);
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(13, 100);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(174, 23);
            this.cancel.TabIndex = 3;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // JobCostingGUI
            // 
            this.AcceptButton = this.analyzebtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel;
            this.ClientSize = new System.Drawing.Size(198, 130);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.closebtn);
            this.Controls.Add(this.analyzebtn);
            this.Controls.Add(this.openbtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "JobCostingGUI";
            this.Text = "Job Costing";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button openbtn;
        private System.Windows.Forms.Button analyzebtn;
        private System.Windows.Forms.Button closebtn;
        private System.Windows.Forms.Button cancel;
    }
}

