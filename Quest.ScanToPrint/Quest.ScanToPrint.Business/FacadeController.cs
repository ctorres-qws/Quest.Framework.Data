using Quest.ScanToPrint.Data;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Business
{
    public class FacadeController
    {
        LocalPersistenceStrategiesFactory LocalPersistenceStrategiesFactory { get; set; }
        OnlinePersistenceStrategiesFactory OnlinePersistenceStrategiesFactory { get; set; }
        public FacadeController(LocalPersistenceStrategiesFactory localPersistenceStrategiesFactory, OnlinePersistenceStrategiesFactory onlinePersistenceStrategiesFactory)
        {
            this.LocalPersistenceStrategiesFactory = localPersistenceStrategiesFactory;
            this.OnlinePersistenceStrategiesFactory = onlinePersistenceStrategiesFactory;
        }
        #region CRUD methods
        public List<JobShippingLabelColor> GetJobShippingLabelColors()
        {
            return LocalPersistenceStrategiesFactory.GetJobShippingLabelColorPersistenceStrategy().GetCollection().ToList();
        }
        public void AddJobShippingLabelColor(JobShippingLabelColor jobShippingLabelColor)
        {
            LocalPersistenceStrategiesFactory.GetJobShippingLabelColorPersistenceStrategy().Add(jobShippingLabelColor);
        }
        public void EditJobShippingLabelColor(JobShippingLabelColor jobShippingLabelColor)
        {
            LocalPersistenceStrategiesFactory.GetJobShippingLabelColorPersistenceStrategy().Edit(jobShippingLabelColor);
        }
        public void DeleteJobShippingLabelColor(JobShippingLabelColor jobShippingLabelColor)
        {
            LocalPersistenceStrategiesFactory.GetJobShippingLabelColorPersistenceStrategy().Delete(jobShippingLabelColor);
        }
        public void AddBarcode(Barcodes barcode)
        {
            LocalPersistenceStrategiesFactory.GetBarcodesPersistanceStrategy().Add(barcode);
        }
        public void EditBarcode(Barcodes barcode)
        {
            LocalPersistenceStrategiesFactory.GetBarcodesPersistanceStrategy().Edit(barcode);
        }
        public void DeleteBarcode(Barcodes barcode)
        {
            LocalPersistenceStrategiesFactory.GetBarcodesPersistanceStrategy().Delete(barcode);
        }
        public List<Barcodes> GetBarcodes()
        {
            return LocalPersistenceStrategiesFactory.GetBarcodesPersistanceStrategy().GetCollection().ToList();
        }
        public void AddGlazing(Glazing glazing)
        {
            OnlinePersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Add(glazing);
        }
        public void EditGlazing(Glazing glazing)
        {
            OnlinePersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Edit(glazing);
        }
        public void DeleteGlazing(Glazing glazing)
        {
            OnlinePersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Delete(glazing);
        }
        public List<Glazing> GetGlazings()
        {
            return new List<Glazing>(OnlinePersistenceStrategiesFactory.GetGlazingPersistenceStrategy().GetCollection());
        }
        public List<Job> GetJobs()
        {
            return OnlinePersistenceStrategiesFactory.GetJobPersistenceStrategy().GetCollection().ToList();
        }
        public void AddRangeJobShippingLabelColor(List<JobShippingLabelColor> setOfItems)
        {
            foreach(JobShippingLabelColor item in setOfItems)
            {
                AddJobShippingLabelColor(item);
            }
        }
        public void EditRangeJobShippingLabelColor(List<JobShippingLabelColor> setOfItems)
        {
            foreach (JobShippingLabelColor item in setOfItems)
            {
                EditJobShippingLabelColor(item);
            }
        }
        public List<RegisteredTag> GetRegisteredTags()
        {
            return new List<RegisteredTag>(LocalPersistenceStrategiesFactory.GetRegisteredTagPersistenceStrategy().GetCollection());
        }
        public void AddRegisteredTag(RegisteredTag registeredTag)
        {
            LocalPersistenceStrategiesFactory.GetRegisteredTagPersistenceStrategy().Add(registeredTag);
        }
        public void AddRangeRegisteredTag(List<RegisteredTag> registeredTags)
        {
            foreach(RegisteredTag registeredTag in registeredTags)
            {
                AddRegisteredTag(registeredTag);
            }
        }
        public List<JobOrderEntryData> GetJobOrderEntryData(string job)
        {
            return new List<JobOrderEntryData>(OnlinePersistenceStrategiesFactory.GetJobOrderEntryDataPersistenceStrategy().GetCollectionFromAJobTable(job).ToList());
        }
        public List<Printer> GetPrintersStatusFromLocalDataSource()
        {
            return new List<Printer>(LocalPersistenceStrategiesFactory.GetPrinterPersistenceStrategy().GetCollection());
        }
        public List<Printer> GetPrintersStatusFromOnlineDataSource()
        {
            return new List<Printer>(OnlinePersistenceStrategiesFactory.GetPrinterPersistenceStrategy().GetCollection());
        }
        public void EditPrinterStatus(Printer printer)
        {
            LocalPersistenceStrategiesFactory.GetPrinterPersistenceStrategy().Edit(printer);
        }
        //public List<JobOrderEntryData> GetJobOrderEntryData(string job)
        //{

        //    return new List<JobOrderEntryData>(OnlineOLEDBPersistenceStrategiesFactory.GetJobOrderEntryDataPersistenceStrategy().GetCollection();
        //}
        #endregion
        /// <summary>
        /// Download new job color matches and update those ones that have been modified
        /// </summary>
        public void UpdateLocalData()
        {
            List<Job> jobs = GetJobs().Distinct().ToList();

            List<JobShippingLabelColor> JobShippingLabelColors = GetJobShippingLabelColors();
            List<JobShippingLabelColor> JobShippingLabelColorsToUpdate = new List<JobShippingLabelColor>();
            List<JobShippingLabelColor> JobShippingLabelColorsToAdd = new List<JobShippingLabelColor>();

            foreach (Job job in jobs)
            {
                JobShippingLabelColor relatedJobShippingLabelColor = JobShippingLabelColors.FirstOrDefault(x => x.Job == job.JOB);
                if(relatedJobShippingLabelColor == null)
                {
                    relatedJobShippingLabelColor = new JobShippingLabelColor()
                    {
                        Color = job.ShippingLabelColor,
                        Job = job.JOB
                    };
                    JobShippingLabelColorsToAdd.Add(relatedJobShippingLabelColor);
                }
                else
                {
                    if(relatedJobShippingLabelColor.Color != job.ShippingLabelColor)
                    {
                        relatedJobShippingLabelColor.Color = job.ShippingLabelColor;
                        JobShippingLabelColorsToUpdate.Add(relatedJobShippingLabelColor);
                    }
                }
            }
            AddRangeJobShippingLabelColor(JobShippingLabelColorsToAdd);
            EditRangeJobShippingLabelColor(JobShippingLabelColorsToUpdate);            
        }
        public void UpdateTagData()
        {
            List<Job> jobs = GetJobs().Distinct().Where(x => !string.IsNullOrEmpty(x.ShippingLabelColor)).ToList();

            List<RegisteredTag> registeredTags = GetRegisteredTags();
            List<RegisteredTag> notRegisteredTags = new List<RegisteredTag>();

            foreach (Job job in jobs)
            {
                List<JobOrderEntryData> jobOrderEntryData;
                try
                {
                    jobOrderEntryData = GetJobOrderEntryData(job.JOB.Trim());
                }
                catch(Exception ex)
                {
                    Log(string.Format("{0} does not have a job table", job.JOB.Trim()), ex);
                    continue;
                }

                foreach(JobOrderEntryData item in jobOrderEntryData)
                {
                    if(!registeredTags.Any(x => x.Job == item.Job && x.Floor == item.Floor && x.Tag == item.Tag))
                    {
                        notRegisteredTags.Add(new RegisteredTag()
                        {
                            Job = item.Job,
                            Tag = item.Tag,
                            Floor = item.Floor,
                            ContainsSW = item.ContainsSW
                        });
                    }
                }
            }
            AddRangeRegisteredTag(notRegisteredTags);
        }
        public void UploadBarcodes()
        {
            List<Barcodes> barcodes = GetBarcodes().Where(x => !x.SentDatabase).ToList();
            //List<Glazing> uploadedBarcodes = GetGlazings();
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            List<Barcodes> barcodesToModify = new List<Barcodes>();
            List<RegisteredTag> registeredTags = GetRegisteredTags();

            foreach (Barcodes barcode in barcodes)
            {
                try
                {
                    string floor = barcode.Barcode.Split('-')[0].Substring(3);
                    RegisteredTag registeredTag = registeredTags.FirstOrDefault(x => x.Job == barcode.Job && x.Floor == floor && x.Tag == barcode.Tag);
                    bool containsSW = registeredTag != null ? registeredTag.ContainsSW : false;

                    int line = containsSW ? 4 : barcode.Line;

                    AddGlazing(new Glazing()
                    {
                        Barcode = barcode.Barcode,
                        Job = barcode.Job,
                        DateTime = barcode.ScanDate,
                        Floor = floor,
                        Day = barcode.ScanDate.Day,
                        Month = barcode.ScanDate.Month,
                        Year = barcode.ScanDate.Year,
                        Dept = "GLAZING",
                        Time = barcode.ScanDate.TimeOfDay,
                        Week = cal.GetWeekOfYear(barcode.ScanDate, dfi.CalendarWeekRule, dfi.FirstDayOfWeek),
                        Tag = barcode.Tag,
                        Count = 1,
                        Employee = string.Format("{0}{0}{0}{0}", line.ToString()),
                        FirstComplete = "TRUE",
                        Joints = 0,
                        O1 = string.Format("{0}{0}{0}{0}", line.ToString()),
                        O2 = "",
                        O3 = "",
                        O4 = "",
                        O5 = "",
                        O6 = "",
                        O7 = "",
                        O8 = "",
                        ONumber = 0,
                        Openings = 0,
                        ScanCount = 1
                    });
                    barcode.SentDatabase = true;

                    barcodesToModify.Add(barcode);
                }
                catch (Exception ex)
                {
                    Log("Barcodes uploading failed", ex);
                }
            }
            if (barcodesToModify.Count > 0)
            {
                foreach(Barcodes barcode in barcodesToModify)
                {
                    EditBarcode(barcode);
                }
            }
        }

        private bool IsBarcodeDataCorrect(BarcodeReading barcode)
        {
            List<RegisteredTag> tags = GetRegisteredTags();

            if (!tags.Any(x => x.Job.Trim().ToUpper() == barcode.Job.Trim().ToUpper()))
                throw new Exception(string.Format("Job {0} is not registered", barcode.Job));

            if (!tags.Any(x => x.Job.Trim().ToUpper() == barcode.Job.Trim().ToUpper() && x.Floor.Trim().ToUpper() == barcode.Floor.Trim().ToUpper()))
                throw new Exception(string.Format("Floor {0} does not belong to {1} job", barcode.Floor, barcode.Job));

            if (!tags.Any(x => x.Job.Trim().ToUpper() == barcode.Job.Trim().ToUpper() && x.Floor.Trim().ToUpper() == barcode.Floor.Trim().ToUpper() && x.Tag.Trim().ToUpper() == barcode.Tag.Trim().ToUpper()))
                throw new Exception(string.Format("Tag {0} does not belong to the Floor {1} from {2} job", barcode.Tag, barcode.Floor, barcode.Job));

            return true;
        }
        private bool DoesThisJobHaveAColorAssigned(string job)
        {
            JobShippingLabelColor color = GetJobShippingLabelColors().FirstOrDefault(x => x.Job == job);

            if (color == null)
                throw new Exception(string.Format("Job {0} is not registered", job));
            if (string.IsNullOrEmpty(color.Color))
                return false;
            return true;
        }
        public BarcodeReading RegisterBarcodeScan(string scannedBarcode)
        {
            BarcodeReading br = ParseBarcode(scannedBarcode);

            if (!DoesThisJobHaveAColorAssigned(br.Job))
                throw new Exception(string.Format("JOB {0} does not have an assigned color", br.Job));

            if (!IsBarcodeDataCorrect(br))
                throw new Exception(string.Format("Job"));

            br.ScanDate = DateTime.Now;

            Barcodes barcode = GetBarcodes().FirstOrDefault(x => x.Barcode == br.Barcode);

            if (barcode == null)
            {
                barcode = new Barcodes()
                {
                    Barcode = br.Barcode,
                    Job = br.Job,
                    Line = br.Line,
                    ScanDate = br.ScanDate,
                    SentDatabase = false,
                    SentPrint = false,
                    Tag = br.Tag
                };
                AddBarcode(barcode);
            }

            return br;
        }
        public string GetJobColor(string job)
        {
            JobShippingLabelColor jobShippingLabelColor = GetJobShippingLabelColors().FirstOrDefault(x => x.Job == job);

            if (jobShippingLabelColor == null)
                return "#ffffff";

            return jobShippingLabelColor.Color;
        }
        public void PrintLabel(Barcodes barcode)
        {
            //TO DO
        }
        private BarcodeReading ParseBarcode(string barcode)
        {
            if (barcode.Count(x => x == '-') < 2)
                throw new Exception("Job, floor and Tag information cannot be decoded");

            if(barcode.Trim().Split('-').Length> 3)
                throw new Exception("Barcode has extra information");

            return new BarcodeReading()
            {
                Barcode = barcode.Substring(3),
                Floor = barcode.Trim().Split('-')[1].Substring(3),
                Job = barcode.Trim().Split('-')[1].Substring(0, 3),
                Line = Convert.ToInt32(barcode.Trim().Substring(1, 1)),
                Tag = string.Format("-{0}", barcode.Trim().Split('-')[2])
            };
        }
        public void Log(string description, Exception ex = null)
        {
            LocalPersistenceStrategiesFactory.GetLogPersistenceStrategy().Add(
                new Log()
                {
                    Description = description,
                    DateTime = DateTime.Now,
                    ExceptionMessage = (ex == null || ex.Message == null) ? "" : ex.Message.ToString(),
                    ExceptionStackTrace = ex == null ? "" : ex.StackTrace.ToString()
                });
        }
        public void UpdatePrintersStatus()
        {
            List<Printer> localData = GetPrintersStatusFromLocalDataSource(), onlineData = GetPrintersStatusFromOnlineDataSource();

            foreach(Printer printerOnline in onlineData)
            {
                Printer printerLocal = localData.First(x => x.GlazingLine == printerOnline.GlazingLine);
                if(printerLocal.Active != printerOnline.Active || printerLocal.BackupPrinter != printerOnline.BackupPrinter)
                {
                    printerLocal.Active = printerOnline.Active;
                    printerLocal.BackupPrinter = printerOnline.BackupPrinter;
                    EditPrinterStatus(printerLocal);
                }
            }
        }
        public int GetTargetPrinter(int line)
        {
            Printer printer = GetPrintersStatusFromLocalDataSource().First(x => x.GlazingLine == line);

            if (printer.Active)
                return printer.GlazingLine;

            return printer.BackupPrinter;
        }
    }
}
