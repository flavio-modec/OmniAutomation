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
            this.unitCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.genTime = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.applyTime = new System.Windows.Forms.Button();
            this.path = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.xmlPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // getReport
            // 
            this.getReport.Location = new System.Drawing.Point(130, 232);
            this.getReport.Name = "getReport";
            this.getReport.Size = new System.Drawing.Size(76, 25);
            this.getReport.TabIndex = 0;
            this.getReport.Text = "Get Reports";
            this.getReport.UseVisualStyleBackColor = true;
            this.getReport.Click += new System.EventHandler(this.getReport_Click);
            // 
            // status
            // 
            this.status.Location = new System.Drawing.Point(44, 118);
            this.status.Name = "status";
            this.status.ReadOnly = true;
            this.status.Size = new System.Drawing.Size(124, 20);
            this.status.TabIndex = 1;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(44, 154);
            this.progressBar1.Maximum = 63;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(124, 20);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 2;
            // 
            // unitCode
            // 
            this.unitCode.AccessibleName = "";
            this.unitCode.Location = new System.Drawing.Point(93, 72);
            this.unitCode.Name = "unitCode";
            this.unitCode.Size = new System.Drawing.Size(38, 20);
            this.unitCode.TabIndex = 3;
            this.unitCode.Tag = "";
            this.unitCode.Text = "10468";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Unit Code:";
            // 
            // genTime
            // 
            this.genTime.Location = new System.Drawing.Point(116, 195);
            this.genTime.Name = "genTime";
            this.genTime.Size = new System.Drawing.Size(38, 20);
            this.genTime.TabIndex = 5;
            this.genTime.Text = "17:15";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 198);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Generation time:";
            // 
            // applyTime
            // 
            this.applyTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.applyTime.Location = new System.Drawing.Point(160, 193);
            this.applyTime.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.applyTime.Name = "applyTime";
            this.applyTime.Size = new System.Drawing.Size(44, 24);
            this.applyTime.TabIndex = 7;
            this.applyTime.Text = "Apply";
            this.applyTime.UseVisualStyleBackColor = true;
            this.applyTime.Click += new System.EventHandler(this.applyTime_Click);
            // 
            // path
            // 
            this.path.Location = new System.Drawing.Point(93, 19);
            this.path.Name = "path";
            this.path.Size = new System.Drawing.Size(90, 20);
            this.path.TabIndex = 8;
            this.path.Text = "D:\\OmniReports";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Reports:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "XML 004:";
            // 
            // xmlPath
            // 
            this.xmlPath.Location = new System.Drawing.Point(93, 46);
            this.xmlPath.Name = "xmlPath";
            this.xmlPath.Size = new System.Drawing.Size(90, 20);
            this.xmlPath.TabIndex = 11;
            this.xmlPath.Text = "D:\\XLR\\Reports";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(226, 271);
            this.ControlBox = false;
            this.Controls.Add(this.xmlPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.path);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.applyTime);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.status);
            this.Controls.Add(this.genTime);
            this.Controls.Add(this.getReport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.unitCode);
            this.Name = "Form1";
            this.Text = "Omni Automation";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button getReport;
        public System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox status;
        private System.Windows.Forms.TextBox unitCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox genTime;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button applyTime;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox path;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox xmlPath;
    }
}

