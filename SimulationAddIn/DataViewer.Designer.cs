namespace SimulationAddIn
{
    partial class DataViewer
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
            this.cbxFilter = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lbxSheetList = new System.Windows.Forms.ListBox();
            this.btnExportFiles = new System.Windows.Forms.Button();
            this.btnDHAMenu = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbxFilter
            // 
            this.cbxFilter.FormattingEnabled = true;
            this.cbxFilter.Location = new System.Drawing.Point(12, 35);
            this.cbxFilter.Name = "cbxFilter";
            this.cbxFilter.Size = new System.Drawing.Size(121, 21);
            this.cbxFilter.TabIndex = 0;
            this.cbxFilter.SelectedIndexChanged += new System.EventHandler(this.cbxFilter_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Worksheet Filter:";
            // 
            // lbxSheetList
            // 
            this.lbxSheetList.FormattingEnabled = true;
            this.lbxSheetList.Location = new System.Drawing.Point(12, 76);
            this.lbxSheetList.Name = "lbxSheetList";
            this.lbxSheetList.Size = new System.Drawing.Size(121, 95);
            this.lbxSheetList.TabIndex = 2;
            this.lbxSheetList.SelectedIndexChanged += new System.EventHandler(this.lbxSheetList_SelectedIndexChanged);
            // 
            // btnExportFiles
            // 
            this.btnExportFiles.Location = new System.Drawing.Point(12, 186);
            this.btnExportFiles.Name = "btnExportFiles";
            this.btnExportFiles.Size = new System.Drawing.Size(121, 23);
            this.btnExportFiles.TabIndex = 3;
            this.btnExportFiles.Text = "Export Files";
            this.btnExportFiles.UseVisualStyleBackColor = true;
            // 
            // btnDHAMenu
            // 
            this.btnDHAMenu.Location = new System.Drawing.Point(12, 229);
            this.btnDHAMenu.Name = "btnDHAMenu";
            this.btnDHAMenu.Size = new System.Drawing.Size(121, 23);
            this.btnDHAMenu.TabIndex = 4;
            this.btnDHAMenu.Text = "DHA Menu";
            this.btnDHAMenu.UseVisualStyleBackColor = true;
            // 
            // DataViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(156, 261);
            this.Controls.Add(this.btnDHAMenu);
            this.Controls.Add(this.btnExportFiles);
            this.Controls.Add(this.lbxSheetList);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbxFilter);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "DataViewer";
            this.Text = "Data Viewer";
            this.Load += new System.EventHandler(this.DataViewer_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxFilter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lbxSheetList;
        private System.Windows.Forms.Button btnExportFiles;
        private System.Windows.Forms.Button btnDHAMenu;
    }
}