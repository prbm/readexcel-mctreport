namespace ReadExcelFile
{
    partial class frmPreviewExcel
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
            this.dataGridSpreadSheetOverview = new System.Windows.Forms.DataGridView();
            this.pbOk = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSpreadSheetOverview)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridSpreadSheetOverview
            // 
            this.dataGridSpreadSheetOverview.AllowUserToAddRows = false;
            this.dataGridSpreadSheetOverview.AllowUserToDeleteRows = false;
            this.dataGridSpreadSheetOverview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridSpreadSheetOverview.Location = new System.Drawing.Point(12, 12);
            this.dataGridSpreadSheetOverview.Name = "dataGridSpreadSheetOverview";
            this.dataGridSpreadSheetOverview.ReadOnly = true;
            this.dataGridSpreadSheetOverview.Size = new System.Drawing.Size(807, 354);
            this.dataGridSpreadSheetOverview.TabIndex = 0;
            // 
            // pbOk
            // 
            this.pbOk.Location = new System.Drawing.Point(376, 375);
            this.pbOk.Name = "pbOk";
            this.pbOk.Size = new System.Drawing.Size(75, 23);
            this.pbOk.TabIndex = 1;
            this.pbOk.Text = "&Ok";
            this.pbOk.UseVisualStyleBackColor = true;
            this.pbOk.Click += new System.EventHandler(this.pbOk_Click);
            // 
            // frmPreviewExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(830, 407);
            this.Controls.Add(this.pbOk);
            this.Controls.Add(this.dataGridSpreadSheetOverview);
            this.Name = "frmPreviewExcel";
            this.Text = "Excel WorkSheet Preview";
            this.Load += new System.EventHandler(this.frmPreviewExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSpreadSheetOverview)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridSpreadSheetOverview;
        private System.Windows.Forms.Button pbOk;
    }
}