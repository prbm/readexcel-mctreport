namespace ReadExcelFile
{
    partial class frmManMonthGenerator
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
            this.label1 = new System.Windows.Forms.Label();
            this.edPMSWorkbook = new System.Windows.Forms.TextBox();
            this.pbBrowsePMSWorkbook = new System.Windows.Forms.Button();
            this.pbGetPMSData = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.pbMySQL = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "PMS &Generated Workbook:";
            // 
            // edPMSWorkbook
            // 
            this.edPMSWorkbook.Location = new System.Drawing.Point(157, 17);
            this.edPMSWorkbook.Name = "edPMSWorkbook";
            this.edPMSWorkbook.Size = new System.Drawing.Size(360, 20);
            this.edPMSWorkbook.TabIndex = 1;
            // 
            // pbBrowsePMSWorkbook
            // 
            this.pbBrowsePMSWorkbook.Location = new System.Drawing.Point(527, 15);
            this.pbBrowsePMSWorkbook.Name = "pbBrowsePMSWorkbook";
            this.pbBrowsePMSWorkbook.Size = new System.Drawing.Size(104, 23);
            this.pbBrowsePMSWorkbook.TabIndex = 2;
            this.pbBrowsePMSWorkbook.Text = "&Browse...";
            this.pbBrowsePMSWorkbook.UseVisualStyleBackColor = true;
            this.pbBrowsePMSWorkbook.Click += new System.EventHandler(this.pbBrowsePMSWorkbook_Click);
            // 
            // pbGetPMSData
            // 
            this.pbGetPMSData.Location = new System.Drawing.Point(527, 44);
            this.pbGetPMSData.Name = "pbGetPMSData";
            this.pbGetPMSData.Size = new System.Drawing.Size(104, 23);
            this.pbGetPMSData.TabIndex = 3;
            this.pbGetPMSData.Text = "&Get PMS Data";
            this.pbGetPMSData.UseVisualStyleBackColor = true;
            this.pbGetPMSData.Click += new System.EventHandler(this.pbGetPMSData_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(58, 134);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(113, 40);
            this.button1.TabIndex = 4;
            this.button1.Text = "PostGres Test";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pbMySQL
            // 
            this.pbMySQL.Location = new System.Drawing.Point(58, 180);
            this.pbMySQL.Name = "pbMySQL";
            this.pbMySQL.Size = new System.Drawing.Size(113, 40);
            this.pbMySQL.TabIndex = 5;
            this.pbMySQL.Text = "MySQL Test";
            this.pbMySQL.UseVisualStyleBackColor = true;
            this.pbMySQL.Click += new System.EventHandler(this.pbMySQL_Click);
            // 
            // frmManMonthGenerator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(644, 266);
            this.Controls.Add(this.pbMySQL);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pbGetPMSData);
            this.Controls.Add(this.pbBrowsePMSWorkbook);
            this.Controls.Add(this.edPMSWorkbook);
            this.Controls.Add(this.label1);
            this.Name = "frmManMonthGenerator";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Man Month Report Creator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox edPMSWorkbook;
        private System.Windows.Forms.Button pbBrowsePMSWorkbook;
        private System.Windows.Forms.Button pbGetPMSData;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button pbMySQL;
    }
}