namespace Quest.ScantToPrint.Presentation
{
    partial class ScanToPrintMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ScanToPrintMain));
            this.txtQRScan = new System.Windows.Forms.TextBox();
            this.pdfViewer = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.pdfViewer)).BeginInit();
            this.SuspendLayout();
            // 
            // txtQRScan
            // 
            this.txtQRScan.Location = new System.Drawing.Point(766, 20);
            this.txtQRScan.Name = "txtQRScan";
            this.txtQRScan.Size = new System.Drawing.Size(100, 20);
            this.txtQRScan.TabIndex = 2;
            this.txtQRScan.TextChanged += new System.EventHandler(this.txtQRScan_TextChanged);
            this.txtQRScan.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtQRScan_KeyUp);
            this.txtQRScan.Leave += new System.EventHandler(this.txtQRScan_Leave);
            // 
            // pdfViewer
            // 
            this.pdfViewer.Enabled = true;
            this.pdfViewer.Location = new System.Drawing.Point(22, 60);
            this.pdfViewer.Name = "pdfViewer";
            this.pdfViewer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("pdfViewer.OcxState")));
            this.pdfViewer.Size = new System.Drawing.Size(844, 374);
            this.pdfViewer.TabIndex = 3;
            // 
            // ScanToPrintMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CausesValidation = false;
            this.ClientSize = new System.Drawing.Size(892, 469);
            this.Controls.Add(this.pdfViewer);
            this.Controls.Add(this.txtQRScan);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "ScanToPrintMain";
            this.Text = "Scan to Print";
            this.TopMost = true;
            this.Deactivate += new System.EventHandler(this.ScanToPrintMain_Deactivate);
            this.Load += new System.EventHandler(this.ScanToPrintMain_Load);
            this.Leave += new System.EventHandler(this.ScanToPrintMain_Leave);
            ((System.ComponentModel.ISupportInitialize)(this.pdfViewer)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtQRScan;
        private AxAcroPDFLib.AxAcroPDF pdfViewer;
    }
}

