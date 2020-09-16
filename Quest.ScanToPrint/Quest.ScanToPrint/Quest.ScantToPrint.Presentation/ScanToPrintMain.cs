using iText.Kernel.Events;
using Microsoft.Win32;
using Quest.ScanToPrint.Business;
using Quest.ScanToPrint.Data;
using Quest.ScanToPrint.FocusHandler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quest.ScantToPrint.Presentation
{
    public partial class ScanToPrintMain : Form
    {
        FacadeController controller;
        BackgroundQueue backgroundQueue;
        static string GetConnectionStringByName(string name)
        {
            // Assume failure.
            string returnValue = null;

            // Look for the name in the connectionStrings section.
            ConnectionStringSettings settings =
                ConfigurationManager.ConnectionStrings[name];

            // If found, return the connection string.
            if (settings != null)
                returnValue = settings.ConnectionString;

            return returnValue;
        }
        public ScanToPrintMain()
        {
            InitializeComponent();
            controller = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(GetConnectionStringByName("localDB")),
                new OnlineOLEDBPersistenceStrategiesFactory(GetConnectionStringByName("onlineDB")));
            backgroundQueue = new BackgroundQueue();
            System.Timers.Timer t = new System.Timers.Timer(TimeSpan.FromMinutes(5).TotalMilliseconds); 
            t.AutoReset = true;
            t.Elapsed += new System.Timers.ElapsedEventHandler(SyncData);
            t.Start();
            //this.ShowDialog();
        }
        private void MaximizeWindow()
        {
            this.WindowState = FormWindowState.Normal;
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
        }
        private void SyncData(object s, System.Timers.ElapsedEventArgs e)
        {
            if (CheckForInternetConnection())
            {
                try
                {
                    controller.UploadBarcodes();
                    controller.UpdateLocalData();
                    controller.Log("Data synchronized successfully");
                }
                catch(Exception ex)
                {
                    controller.Log(ex.Message, ex);
                }
                
            }
            else
            {
                controller.Log("NO NETWORK CONNECTION: Data could not be synchronized");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void txtQRScan_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtQRScan_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (char)Keys.Return)
            {
                e.Handled = true;
                
                PDFcreator pDFcreator = new PDFcreator();
                
                this.txtQRScan.LostFocus += new System.EventHandler(this.returnFocus);

                pDFcreator.CreatePDF6x4(new BarcodeReading() { Barcode = "", Floor = "", Job = "", Tag = "" }, "#000000");

                pdfViewer.setShowToolbar(false);
                pdfViewer.setShowScrollbars(true);
                pdfViewer.LoadFile(@"..\..\Labels\label.pdf");
                try
                {
                    BarcodeReading barcodeReading = controller.RegisterBarcodeScan(txtQRScan.Text);
                    pDFcreator.CreatePDF6x4(barcodeReading, controller.getJobColor(barcodeReading.Job));
                    pdfViewer.LoadFile(string.Format(@"..\..\Labels\{0}.pdf", barcodeReading.Barcode));

                    backgroundQueue.QueueTask(() => Print(barcodeReading.Barcode, barcodeReading.Line));
                }
                catch (Exception ex)
                {
                    controller.Log("Printing failed.", ex);
                    pdfViewer.LoadFile("none");
                }
                
                
                
                //pdfViewer.printAll();
                txtQRScan.Clear();
            }
        }
        private void Print(string barcode, int glazingLine)
        {
            string printer = "";

            switch (glazingLine)
            {
                case 1:
                    printer = ConfigurationManager.AppSettings["Printer_GL1"].ToString();
                    break;
                case 2:
                    printer = ConfigurationManager.AppSettings["Printer_GL2"].ToString();
                    break;
                case 3:
                    printer = ConfigurationManager.AppSettings["Printer_GL3"].ToString();
                    break;
                default:
                    printer = ConfigurationManager.AppSettings["Printer_Default"].ToString();
                    break;
            }

            PrintPDF(string.Format(@"{0}\{1}.pdf", ConfigurationManager.AppSettings["labelsDirectory"].ToString(), barcode), printer);
        }
        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                using (client.OpenRead("http://google.com/generate_204"))
                    return true;
            }
            catch
            {
                return false;
            }
        }
        private void returnFocus(object sender, EventArgs e)
        {
            txtQRScan.Focus();
        }

        private void ScanToPrintMain_Load(object sender, EventArgs e)
        {

        }
        public Boolean PrintPDF(string pdfFileName, string printer)
        {
            try
            {
                var exePath = ConfigurationManager.AppSettings["SumatraEXE"].ToString();

                var args = $"-print-to {printer} {pdfFileName}";
                controller.Log(args);
                var process = Process.Start(exePath, args);

                

                process.WaitForExit();
                return true;
            }
            catch
            {
                return false;
            }
        }
        private void SetFocusOnMainWindown()
        {
            FocusHandler.SetFocus(string.Format("{0}.exe", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name));
        }
        private void txtQRScan_Leave(object sender, EventArgs e)
        {
            txtQRScan.Focus();
        }

        private void ScanToPrintMain_Leave(object sender, EventArgs e)
        {
            //SetFocusOnMainWindown();
        }

        private void ScanToPrintMain_Deactivate(object sender, EventArgs e)
        {
            //SetFocusOnMainWindown();
        }
    }
}
