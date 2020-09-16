using Quest.Framework.Data.Entities;
using Quest.Framework.Persistance;
using Quest.Framework.Persistance.OLEDB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Data.Persistence.OLEDB
{
    public class JobOLEDBPersistenceStrategy : OLEDBPersistenceStrategy<Job>
    {
        public JobOLEDBPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineEditCommand(ref OleDbCommand command, Job entity)
        {
            command.CommandText = string.Format("UPDATE Z_Jobs SET ShippingLabelColor = '{0}' WHERE JOB = '{1}'", entity.ShippingLabelColor, entity.JOB);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            string query = "SELECT [ID], 	[JOB], 	        [Job_Name], 	[Job_Address], 	    [Job_City], 	    [Parent], 	        [Floors], 	    [Material],";	
                query += "[FrameStyle], 	[Sill], 	    [Ext_Colour], 	[Int_Colour], 	    [Manager], 	        [ManagerEmail], 	[Completed], 	[ExtGlass],";
                query += "[IntGlass], 	    [ExtGlassDoor], [IntGlassDoor], [FixWindowThick], 	[SwingDoorThick], 	[CasAwnThick], 	    [SUSpacer], 	";
                query += "[OVSpacer], 	    [SWSpacer], 	[SillType], 	[SpacerColour], 	[AwnStyle], 	    [LouverStyle], 	    [GlassFlush], 	[PanelPunch],";
                query += "[R3VentSize], 	[Job_Country], 	[Engineer], 	[EngineerEmail], 	[RecipientList], 	[ImporterRecord], 	[ImporterAddress], 	";
                query += "[ImporterTaxID], 	[ExporterTaxID],[ShiftWindow], 	[BeautyStyle], 	    [DoorGlassStop], 	[WindowGlassStop], 	[ColorMatch], 	[MaxHoist],";
                query += "[VStockLength], 	[JobStatus], 	[ModifiedDate], [OnSiteDate], 	    [StopColor], 	    [LouverStop], 	    [NoDoors], 	    [NoAwnings],";
                query += "[Screen],         [SDSill], 	    [SWSill], 	    [ShiftVentSize],    [ShippingLabelColor] FROM Z_Jobs";

            command.CommandText = query;
        }
        protected override Job SelectResultItemMapping(DataRow dataRow)
        {
            Job entity = new Job();
            entity.ID = Convert.ToInt32(dataRow["ID"]); 
            entity.JOB = dataRow["JOB"].ToString();
            entity.ShippingLabelColor = dataRow["ShippingLabelColor"].ToString();
            entity.Job_Name = dataRow["Job_Name"].ToString(); 
            entity.Job_Address = dataRow["Job_Address"].ToString(); 
            entity.Job_City = dataRow["Job_City"].ToString(); 
            entity.Parent = dataRow["Parent"].ToString(); 
            entity.Floors = Convert.ToInt32(dataRow["Floors"]); 
            entity.Material = dataRow["Material"].ToString(); 
            entity.FrameStyle = dataRow["FrameStyle"].ToString(); 
            entity.Sill = dataRow["Sill"].ToString(); 
            entity.Ext_Colour = dataRow["Ext_Colour"].ToString(); 
            entity.Int_Colour = dataRow["Int_Colour"].ToString(); 
            entity.Manager = dataRow["Manager"].ToString(); 
            entity.ManagerEmail = dataRow["ManagerEmail"].ToString(); 
            entity.Completed = bool.Parse(dataRow["Completed"].ToString()); 
            entity.ExtGlass = dataRow["ExtGlass"].ToString(); 
            entity.IntGlass = dataRow["IntGlass"].ToString(); 
            entity.ExtGlassDoor = dataRow["ExtGlassDoor"].ToString(); 
            entity.IntGlassDoor = dataRow["IntGlassDoor"].ToString();

            if (string.IsNullOrEmpty(dataRow["FixWindowThick"].ToString()))
                entity.FixWindowThick = null;
            else
                entity.FixWindowThick = Convert.ToInt32(dataRow["FixWindowThick"].ToString());

            if (string.IsNullOrEmpty(dataRow["SwingDoorThick"].ToString()))
                entity.SwingDoorThick = null;
            else
                entity.SwingDoorThick = Convert.ToInt32(dataRow["SwingDoorThick"].ToString());

            if (string.IsNullOrEmpty(dataRow["CasAwnThick"].ToString()))
                entity.CasAwnThick = null;
            else
                entity.CasAwnThick = Convert.ToInt32(dataRow["CasAwnThick"].ToString());

            if (string.IsNullOrEmpty(dataRow["SUSpacer"].ToString()))
                entity.SUSpacer = null;
            else
                entity.SUSpacer = Convert.ToInt32(dataRow["SUSpacer"].ToString());

            if (string.IsNullOrEmpty(dataRow["OVSpacer"].ToString()))
                entity.OVSpacer = null;
            else
                entity.OVSpacer = Convert.ToInt32(dataRow["OVSpacer"].ToString());

            if (string.IsNullOrEmpty(dataRow["SWSpacer"].ToString()))
                entity.SWSpacer = null;
            else
                entity.SWSpacer = Convert.ToInt32(dataRow["SWSpacer"].ToString());

            entity.SillType = dataRow["SillType"].ToString(); 
            entity.SpacerColour = dataRow["SpacerColour"].ToString(); 
            entity.AwnStyle = dataRow["AwnStyle"].ToString(); 
            entity.LouverStyle = dataRow["LouverStyle"].ToString(); 
            entity.GlassFlush = dataRow["GlassFlush"].ToString(); 
            entity.PanelPunch = dataRow["PanelPunch"].ToString();

            if (string.IsNullOrEmpty(dataRow["R3VentSize"].ToString()))
                entity.R3VentSize = null;
            else
                entity.R3VentSize = Convert.ToDecimal(dataRow["R3VentSize"].ToString()); 

            entity.Job_Country = dataRow["Job_Country"].ToString(); 
            entity.Engineer = dataRow["Engineer"].ToString(); 
            entity.EngineerEmail = dataRow["EngineerEmail"].ToString(); 
            entity.RecipientList = dataRow["RecipientList"].ToString(); 
            entity.ImporterRecord = dataRow["ImporterRecord"].ToString(); 
            entity.ImporterAddress = dataRow["ImporterAddress"].ToString(); 
            entity.ImporterTaxID = dataRow["ImporterTaxID"].ToString(); 
            entity.ExporterTaxID = dataRow["ExporterTaxID"].ToString(); 
            entity.ShiftWindow = dataRow["ShiftWindow"].ToString(); 
            entity.BeautyStyle = dataRow["BeautyStyle"].ToString(); 
            entity.DoorGlassStop = dataRow["DoorGlassStop"].ToString(); 
            entity.WindowGlassStop = dataRow["WindowGlassStop"].ToString(); 
            entity.ColorMatch = dataRow["ColorMatch"].ToString();

            if (string.IsNullOrEmpty(dataRow["MaxHoist"].ToString()))
                entity.MaxHoist = null;
            else
                entity.MaxHoist = Convert.ToDecimal(dataRow["MaxHoist"].ToString());

            if (string.IsNullOrEmpty(dataRow["VStockLength"].ToString()))
                entity.VStockLength = null;
            else
                entity.VStockLength = Convert.ToDecimal(dataRow["VStockLength"].ToString());

            entity.JobStatus = dataRow["JobStatus"].ToString();

            if (string.IsNullOrEmpty(dataRow["ModifiedDate"].ToString()))
                entity.ModifiedDate = null;
            else
                entity.ModifiedDate = DateTime.Parse(dataRow["ModifiedDate"].ToString());

            if (string.IsNullOrEmpty(dataRow["OnSiteDate"].ToString()))
                entity.OnSiteDate = null;
            else
                entity.OnSiteDate = DateTime.Parse(dataRow["OnSiteDate"].ToString());
                       
            entity.StopColor = dataRow["StopColor"].ToString();

            entity.LouverStop = dataRow["LouverStop"].ToString(); 
            entity.NoDoors = bool.Parse(dataRow["NoDoors"].ToString());
            entity.NoAwnings = bool.Parse(dataRow["NoAwnings"].ToString()); 
            entity.Screen = dataRow["Screen"].ToString(); 
            entity.SDSill = dataRow["SDSill"].ToString();
            entity.SWSill = dataRow["SWSill"].ToString();

            if (string.IsNullOrEmpty(dataRow["ShiftVentSize"].ToString()))
                entity.ShiftVentSize = null;
            else
                entity.ShiftVentSize = Convert.ToDecimal(dataRow["ShiftVentSize"].ToString());

            return entity;
        }
    }
}
