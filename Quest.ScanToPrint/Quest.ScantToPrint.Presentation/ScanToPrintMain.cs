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
        Dictionary<int, DateTime?> lastScanings;
        int WaitingTime = 30;
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
            int dataSyncFrequency = 6, focusRestorationFrequency = 16, jobOrderEntryDataSyncFrequency = 61;

            if (ConfigurationManager.AppSettings["DataSyncFrequency"] != null)
                int.TryParse(ConfigurationManager.AppSettings["DataSyncFrequency"], out dataSyncFrequency);

            if (ConfigurationManager.AppSettings["FocusRestorationFrequency"] != null)
                int.TryParse(ConfigurationManager.AppSettings["FocusRestorationFrequency"], out focusRestorationFrequency);

            if (ConfigurationManager.AppSettings["JobOrderEntryDataSyncFrequency"] != null)
                int.TryParse(ConfigurationManager.AppSettings["JobOrderEntryDataSyncFrequency"], out jobOrderEntryDataSyncFrequency);

            backgroundQueue = new BackgroundQueue();
            System.Timers.Timer t = new System.Timers.Timer(TimeSpan.FromMinutes(dataSyncFrequency).TotalMilliseconds);
            t.AutoReset = true;
            t.Elapsed += new System.Timers.ElapsedEventHandler(SyncData);
            t.Start();

            System.Timers.Timer t_focus = new System.Timers.Timer(TimeSpan.FromSeconds(focusRestorationFrequency).TotalMilliseconds);
            t_focus.AutoReset = true;
            t_focus.Elapsed += new System.Timers.ElapsedEventHandler(SetFocusOnMainWindown);
            t_focus.Start();

            System.Timers.Timer t_JobOrderEntryData = new System.Timers.Timer(TimeSpan.FromMinutes(jobOrderEntryDataSyncFrequency).TotalMilliseconds);
            t_JobOrderEntryData.AutoReset = true;
            t_JobOrderEntryData.Elapsed += new System.Timers.ElapsedEventHandler(UpdateJobOrderEntryData);
            t_JobOrderEntryData.Start();

            lastScanings = InitializeScanTimes();
            int.TryParse(ConfigurationManager.AppSettings["WaitingPeriod"], out WaitingTime);
        }
        private Dictionary<int, DateTime?> InitializeScanTimes()
        {
            return new Dictionary<int, DateTime?>
            {
                {
                    1, null
                },
                {
                    2, null
                },
                {
                    3, null
                }
            };
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
                    controller.UpdatePrintersStatus();
                    controller.Log("Data synchronized successfully");
                }
                catch (Exception ex)
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
            if (e.KeyValue == (char)Keys.Return && txtQRScan.Text.Trim().Length > 0)
            {
                try
                {
                    e.Handled = true;

                    PDFcreator pDFcreator = new PDFcreator();

                    this.txtQRScan.LostFocus += new System.EventHandler(this.returnFocus);

                    pdfViewer.LoadFile("none");


                    int line = Convert.ToInt32(txtQRScan.Text.Trim().Substring(1, 1));

                    if (line > 3) throw new Exception("Line not registered");

                    if (!(lastScanings[line] != null && DateTime.Now < ((DateTime)lastScanings[line]).AddSeconds(WaitingTime)))
                    {

                        if (txtQRScan.Text.Contains("ScanToPrintTestLabel"))
                        {
                            lastScanings[line] = DateTime.Now;
                            int printer = controller.GetTargetPrinter(line);

                            BarcodeReading barcodeTest = GenerateTestBarcode(printer);
                            
                            if(printer == line)
                                pDFcreator.CreatePDF6x4(GenerateTestBarcode(line), "#dedede");
                            else
                                pDFcreator.CreatePDF6x4FromAnotherLine(GenerateTestBarcode(line), "#dedede");

                            pdfViewer.LoadFile(string.Format(@"..\..\Labels\{0}.pdf", barcodeTest.Barcode));

                            backgroundQueue.QueueTask(() => Print(barcodeTest.Barcode, printer));

                        }
                        else
                        {
                            BarcodeReading barcodeReading = controller.RegisterBarcodeScan(txtQRScan.Text);

                            if (barcodeReading != null) lastScanings[barcodeReading.Line] = DateTime.Now;

                            int printer = controller.GetTargetPrinter(barcodeReading.Line);

                            if (printer == barcodeReading.Line)
                                pDFcreator.CreatePDF6x4(barcodeReading, controller.GetJobColor(barcodeReading.Job));
                            else
                                pDFcreator.CreatePDF6x4FromAnotherLine(barcodeReading, controller.GetJobColor(barcodeReading.Job));

                            pdfViewer.LoadFile(string.Format(@"..\..\Labels\{0}.pdf", barcodeReading.Barcode));

                            backgroundQueue.QueueTask(() => Print(barcodeReading.Barcode, printer));
                        }

                    }
                }
                catch (Exception ex)
                {
                    controller.Log("Printing failed.", ex);
                    pdfViewer.LoadFile("none");
                }
                txtQRScan.Clear();
            }
        }
        private BarcodeReading GenerateTestBarcode(int line)
        {
            return new BarcodeReading()
            {
                Barcode = "ScanToPrintTestLabel",
                Job = "TEST",
                Floor = "TEST",
                Line = line,
                ScanDate = DateTime.Now,
                Tag = "TEST"
            };
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
        private void SetFocusOnMainWindown(object s, System.Timers.ElapsedEventArgs e)
        {
            FocusHandler.SetFocus(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
        }
        private void UpdateJobOrderEntryData(object s, System.Timers.ElapsedEventArgs e)
        {
            if (CheckForInternetConnection())
            {
                try
                {
                    controller.UpdateTagData();
                    controller.Log("Tag data synchronized successfully");
                }
                catch (Exception ex)
                {
                    controller.Log(ex.Message, ex);
                }

            }
            else
            {
                controller.Log("NO NETWORK CONNECTION: Tag data could not be synchronized");
            }
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
