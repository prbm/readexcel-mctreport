namespace ReadExcelFile
{
    partial class frmReadExcelMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReadExcelMain));
            this.tbFileOpenPath = new System.Windows.Forms.TextBox();
            this.lbFileToOpen = new System.Windows.Forms.Label();
            this.btBrowseFile = new System.Windows.Forms.Button();
            this.lbFileDestination = new System.Windows.Forms.Label();
            this.tbFileDestination = new System.Windows.Forms.TextBox();
            this.btProcess = new System.Windows.Forms.Button();
            this.stStripProcessing = new System.Windows.Forms.StatusStrip();
            this.btFileDestination = new System.Windows.Forms.Button();
            this.lbWorkSheets = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lbColumnHeads = new System.Windows.Forms.ListBox();
            this.pbSelectedColumns = new System.Windows.Forms.Button();
            this.pbWSheetPreview = new System.Windows.Forms.Button();
            this.pbGenManMonthReport = new System.Windows.Forms.Button();
            this.pbHRDelayCalculus = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.pbProjectStatusList = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbFileOpenPath
            // 
            resources.ApplyResources(this.tbFileOpenPath, "tbFileOpenPath");
            this.tbFileOpenPath.Name = "tbFileOpenPath";
            // 
            // lbFileToOpen
            // 
            resources.ApplyResources(this.lbFileToOpen, "lbFileToOpen");
            this.lbFileToOpen.Name = "lbFileToOpen";
            // 
            // btBrowseFile
            // 
            resources.ApplyResources(this.btBrowseFile, "btBrowseFile");
            this.btBrowseFile.Name = "btBrowseFile";
            this.btBrowseFile.UseVisualStyleBackColor = true;
            this.btBrowseFile.Click += new System.EventHandler(this.btBrowseFile_Click);
            // 
            // lbFileDestination
            // 
            resources.ApplyResources(this.lbFileDestination, "lbFileDestination");
            this.lbFileDestination.Name = "lbFileDestination";
            // 
            // tbFileDestination
            // 
            resources.ApplyResources(this.tbFileDestination, "tbFileDestination");
            this.tbFileDestination.Name = "tbFileDestination";
            // 
            // btProcess
            // 
            resources.ApplyResources(this.btProcess, "btProcess");
            this.btProcess.Name = "btProcess";
            this.btProcess.UseVisualStyleBackColor = true;
            this.btProcess.Click += new System.EventHandler(this.btProcess_Click);
            // 
            // stStripProcessing
            // 
            resources.ApplyResources(this.stStripProcessing, "stStripProcessing");
            this.stStripProcessing.Name = "stStripProcessing";
            // 
            // btFileDestination
            // 
            resources.ApplyResources(this.btFileDestination, "btFileDestination");
            this.btFileDestination.Name = "btFileDestination";
            this.btFileDestination.UseVisualStyleBackColor = true;
            this.btFileDestination.Click += new System.EventHandler(this.btFileDestination_Click);
            // 
            // lbWorkSheets
            // 
            this.lbWorkSheets.FormattingEnabled = true;
            resources.ApplyResources(this.lbWorkSheets, "lbWorkSheets");
            this.lbWorkSheets.Name = "lbWorkSheets";
            this.lbWorkSheets.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // lbColumnHeads
            // 
            this.lbColumnHeads.FormattingEnabled = true;
            resources.ApplyResources(this.lbColumnHeads, "lbColumnHeads");
            this.lbColumnHeads.MultiColumn = true;
            this.lbColumnHeads.Name = "lbColumnHeads";
            this.lbColumnHeads.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            // 
            // pbSelectedColumns
            // 
            resources.ApplyResources(this.pbSelectedColumns, "pbSelectedColumns");
            this.pbSelectedColumns.Name = "pbSelectedColumns";
            this.pbSelectedColumns.UseVisualStyleBackColor = true;
            this.pbSelectedColumns.Click += new System.EventHandler(this.pbSelectedColumns_Click);
            // 
            // pbWSheetPreview
            // 
            resources.ApplyResources(this.pbWSheetPreview, "pbWSheetPreview");
            this.pbWSheetPreview.Name = "pbWSheetPreview";
            this.pbWSheetPreview.UseVisualStyleBackColor = true;
            this.pbWSheetPreview.Click += new System.EventHandler(this.pbWSheetPreview_Click);
            // 
            // pbGenManMonthReport
            // 
            resources.ApplyResources(this.pbGenManMonthReport, "pbGenManMonthReport");
            this.pbGenManMonthReport.Name = "pbGenManMonthReport";
            this.pbGenManMonthReport.UseVisualStyleBackColor = true;
            this.pbGenManMonthReport.Click += new System.EventHandler(this.pbGenManMonthReport_Click);
            // 
            // pbHRDelayCalculus
            // 
            resources.ApplyResources(this.pbHRDelayCalculus, "pbHRDelayCalculus");
            this.pbHRDelayCalculus.Name = "pbHRDelayCalculus";
            this.pbHRDelayCalculus.UseVisualStyleBackColor = true;
            this.pbHRDelayCalculus.Click += new System.EventHandler(this.pbHRDelayCalculus_Click);
            // 
            // button2
            // 
            resources.ApplyResources(this.button2, "button2");
            this.button2.Name = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // pbProjectStatusList
            // 
            resources.ApplyResources(this.pbProjectStatusList, "pbProjectStatusList");
            this.pbProjectStatusList.Name = "pbProjectStatusList";
            this.pbProjectStatusList.UseVisualStyleBackColor = true;
            this.pbProjectStatusList.Click += new System.EventHandler(this.pbProjectStatusList_Click);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmReadExcelMain
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pbProjectStatusList);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.pbHRDelayCalculus);
            this.Controls.Add(this.pbGenManMonthReport);
            this.Controls.Add(this.pbWSheetPreview);
            this.Controls.Add(this.pbSelectedColumns);
            this.Controls.Add(this.lbColumnHeads);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbWorkSheets);
            this.Controls.Add(this.btFileDestination);
            this.Controls.Add(this.stStripProcessing);
            this.Controls.Add(this.btProcess);
            this.Controls.Add(this.tbFileDestination);
            this.Controls.Add(this.lbFileDestination);
            this.Controls.Add(this.btBrowseFile);
            this.Controls.Add(this.lbFileToOpen);
            this.Controls.Add(this.tbFileOpenPath);
            this.Name = "frmReadExcelMain";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbFileOpenPath;
        private System.Windows.Forms.Label lbFileToOpen;
        private System.Windows.Forms.Button btBrowseFile;
        private System.Windows.Forms.Label lbFileDestination;
        private System.Windows.Forms.TextBox tbFileDestination;
        private System.Windows.Forms.Button btProcess;
        private System.Windows.Forms.StatusStrip stStripProcessing;
        private System.Windows.Forms.Button btFileDestination;
        private System.Windows.Forms.ListBox lbWorkSheets;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lbColumnHeads;
        private System.Windows.Forms.Button pbSelectedColumns;
        private System.Windows.Forms.Button pbWSheetPreview;
        private System.Windows.Forms.Button pbGenManMonthReport;
        private System.Windows.Forms.Button pbHRDelayCalculus;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button pbProjectStatusList;
        private System.Windows.Forms.Button button1;
    }
}

