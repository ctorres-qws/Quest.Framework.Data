using Quest.ScanToPrint.Data;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Business
{
    public class FacadeController
    {
        LocalPersistenceStrategiesFactory LocalPersistenceStrategiesFactory { get; set; }
        OnlineOLEDBPersistenceStrategiesFactory OnlineOLEDBPersistenceStrategiesFactory { get; set; }
        public FacadeController(LocalPersistenceStrategiesFactory localPersistenceStrategiesFactory, OnlineOLEDBPersistenceStrategiesFactory onlineOLEDBPersistenceStrategiesFactory)
        {
            this.LocalPersistenceStrategiesFactory = localPersistenceStrategiesFactory;
            this.OnlineOLEDBPersistenceStrategiesFactory = onlineOLEDBPersistenceStrategiesFactory;
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
            OnlineOLEDBPersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Add(glazing);
        }
        public void EditGlazing(Glazing glazing)
        {
            OnlineOLEDBPersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Edit(glazing);
        }
        public void DeleteGlazing(Glazing glazing)
        {
            OnlineOLEDBPersistenceStrategiesFactory.GetGlazingPersistenceStrategy().Delete(glazing);
        }
        public List<Glazing> GetGlazings()
        {
            return new List<Glazing>(OnlineOLEDBPersistenceStrategiesFactory.GetGlazingPersistenceStrategy().GetCollection());
        }
        public List<Job> GetJobs()
        {
            return OnlineOLEDBPersistenceStrategiesFactory.GetJobPersistenceStrategy().GetCollection().ToList();
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

        public void UploadBarcodes()
        {
            List<Barcodes> barcodes = GetBarcodes().Where(x => !x.SentDatabase).ToList();
            //List<Glazing> uploadedBarcodes = GetGlazings();
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            List<Barcodes> barcodesToModify = new List<Barcodes>();
            foreach (Barcodes barcode in barcodes)
            {
                try
                {
                    AddGlazing(new Glazing()
                    {
                        Barcode = barcode.Barcode,
                        Job = barcode.Job,
                        DateTime = barcode.ScanDate,
                        Floor = barcode.Barcode.Split('-')[0].Substring(3),
                        Day = barcode.ScanDate.Day,
                        Month = barcode.ScanDate.Month,
                        Year = barcode.ScanDate.Year,
                        Dept = "GLAZING",
                        Time = barcode.ScanDate.TimeOfDay,
                        Week = cal.GetWeekOfYear(barcode.ScanDate, dfi.CalendarWeekRule, dfi.FirstDayOfWeek),
                        Tag = barcode.Tag,
                        Count = 1,
                        Employee = string.Format("{0}{0}{0}{0}", barcode.Line.ToString()),
                        FirstComplete = "TRUE",
                        Joints = 0,
                        O1 = string.Format("{0}{0}{0}{0}", barcode.Line.ToString()),
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
                    barcode.Line++;
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
        private bool IsBarcodeInfoCorrect(BarcodeReading barcode)
        {
            List<JobShippingLabelColor> existingJobColors = GetJobShippingLabelColors();

            existingJobColors.Any(x => x.Job == barcode.Job && !string.IsNullOrEmpty(x.Color));



            return true;
        }
        public BarcodeReading RegisterBarcodeScan(string scannedBarcode)
        {
            BarcodeReading br = ParseBarcode(scannedBarcode);

            if (!IsBarcodeInfoCorrect(br))
                throw new Exception(string.Format("JOB {0} does not have an assigned color", br.Job));

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
        public string getJobColor(string job)
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
    }
}
