namespace OmniAutomation
{
    partial class Form1
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
            this.getReport = new System.Windows.Forms.Button();
            this.status = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.apply = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.genTime = new System.Windows.Forms.TextBox();
            this.xmlPath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.repPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.unitCode = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // getReport
            // 
            this.getReport.Location = new System.Drawing.Point(77, 130);
            this.getReport.Name = "getReport";
            this.getReport.Size = new System.Drawing.Size(76, 25);
            this.getReport.TabIndex = 0;
            this.getReport.Text = "Get Reports";
            this.getReport.UseVisualStyleBackColor = true;
            this.getReport.Click += new System.EventHandler(this.getReport_Click);
            // 
            // status
            // 
            this.status.Location = new System.Drawing.Point(49, 38);
            this.status.Name = "status";
            this.status.ReadOnly = true;
            this.status.Size = new System.Drawing.Size(124, 20);
            this.status.TabIndex = 1;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(49, 74);
            this.progressBar1.Maximum = 63;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(124, 20);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(226, 214);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Transparent;
            this.tabPage1.Controls.Add(this.getReport);
            this.tabPage1.Controls.Add(this.progressBar1);
            this.tabPage1.Controls.Add(this.status);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(218, 188);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Run";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.Transparent;
            this.tabPage2.Controls.Add(this.apply);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.genTime);
            this.tabPage2.Controls.Add(this.xmlPath);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.repPath);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.unitCode);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(218, 188);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Options";
            // 
            // apply
            // 
            this.apply.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.apply.Location = new System.Drawing.Point(86, 141);
            this.apply.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.apply.Name = "apply";
            this.apply.Size = new System.Drawing.Size(57, 27);
            this.apply.TabIndex = 29;
            this.apply.Text = "Apply";
            this.apply.UseVisualStyleBackColor = true;
            this.apply.Click += new System.EventHandler(this.apply_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 28;
            this.label2.Text = "Generation time:";
            // 
            // genTime
            // 
            this.genTime.Location = new System.Drawing.Point(105, 101);
            this.genTime.Name = "genTime";
            this.genTime.Size = new System.Drawing.Size(38, 20);
            this.genTime.TabIndex = 27;
            this.genTime.Text = "17:15";
            // 
            // xmlPath
            // 
            this.xmlPath.Location = new System.Drawing.Point(105, 49);
            this.xmlPath.Name = "xmlPath";
            this.xmlPath.Size = new System.Drawing.Size(90, 20);
            this.xmlPath.TabIndex = 26;
            this.xmlPath.Text = "D:\\XLR\\Reports";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(46, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 25;
            this.label4.Text = "XML 004:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(52, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Reports:";
            // 
            // repPath
            // 
            this.repPath.Location = new System.Drawing.Point(105, 22);
            this.repPath.Name = "repPath";
            this.repPath.Size = new System.Drawing.Size(90, 20);
            this.repPath.TabIndex = 23;
            this.repPath.Text = "D:\\OmniReports";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(42, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Unit Code:";
            // 
            // unitCode
            // 
            this.unitCode.AccessibleName = "";
            this.unitCode.Location = new System.Drawing.Point(105, 75);
            this.unitCode.Name = "unitCode";
            this.unitCode.Size = new System.Drawing.Size(38, 20);
            this.unitCode.TabIndex = 21;
            this.unitCode.Tag = "";
            this.unitCode.Text = "10468";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(253, 240);
            this.ControlBox = false;
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Omni Automation";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button getReport;
        public System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox status;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button apply;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox genTime;
        private System.Windows.Forms.TextBox xmlPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox repPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox unitCode;
    }
}

