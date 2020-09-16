using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data.Entities
{
    public class Job
    {
        public int ID { get; set; }
        public string JOB { get; set; }
        public string Job_Name { get; set; }
        public string Job_Address { get; set; }
        public string Job_City { get; set; }
        public string Parent { get; set; }
        public int Floors { get; set; }
        public string Material { get; set; }
        public string FrameStyle { get; set; }
        public string Sill { get; set; }
        public string Ext_Colour { get; set; }
        public string Int_Colour { get; set; }
        public string Manager { get; set; }
        public string ManagerEmail { get; set; }
        public bool Completed { get; set; }
        public string ExtGlass { get; set; }
        public string IntGlass { get; set; }
        public string ExtGlassDoor { get; set; }
        public string IntGlassDoor { get; set; }
        public int FixWindowThick { get; set; }
        public int SwingDoorThick { get; set; }
        public int CasAwnThick { get; set; }
        public int SUSpacer { get; set; }
        public int OVSpacer { get; set; }
        public int SWSpacer { get; set; }
        public string SillType { get; set; }
        public string SpacerColour { get; set; }
        public string AwnStyle { get; set; }
        public string LouverStyle { get; set; }
        public string GlassFlush { get; set; }
        public string PanelPunch { get; set; }
        public int R3VentSize { get; set; }
        public string Job_Country { get; set; }
        public string Engineer { get; set; }
        public string EngineerEmail { get; set; }
        public string RecipientList { get; set; }
        public string ImporterRecord { get; set; }
        public string ImporterAddress { get; set; }
        public string ImporterTaxID { get; set; }
        public string ExporterTaxID { get; set; }
        public string ShiftWindow { get; set; }
        public string BeautyStyle { get; set; }
        public string DoorGlassStop { get; set; }
        public string WindowGlassStop { get; set; }
        public string ColorMatch { get; set; }
        public int MaxHoist { get; set; }
        public int VStockLength { get; set; }
        public string JobStatus { get; set; }
        public DateTime ModifiedDate { get; set; }
        public DateTime OnSiteDate { get; set; }
        public int StopColor { get; set; }
        public string LouverStop { get; set; }
        public bool NoDoors { get; set; }
        public bool NoAwnings { get; set; }
        public string Screen { get; set; }
        public string SDSill { get; set; }
        public string SWSill { get; set; }
        public int ShiftVentSize { get; set; }
        public string ShippingLabelColor { get; set; }
    }
}
