using Quest.Framework.Data.Entities;
using Quest.Framework.Data.Persistence;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.QuestTools.Business
{
    public class FacadeController
    {
        PersistenceStrategiesFactory PersistenceStrategiesFactory { get; set; }
        public FacadeController(PersistenceStrategiesFactory persistenceStrategiesFactory)
        {
            this.PersistenceStrategiesFactory = persistenceStrategiesFactory;
        }
        #region CRUD Methods

        public List<Job> GetJobs()
        {
            return new List<Job>(PersistenceStrategiesFactory.GetJobPersistenceStrategy().GetCollection());
        }
        public List<ShippingColor> GetShippingColors()
        {
            return new List<ShippingColor>(PersistenceStrategiesFactory.GetShippingColorPersistenceStrategy().GetCollection());
        }
        public List<JobShippingColor> GetJobShippingColors()
        {
            return new List<JobShippingColor>(PersistenceStrategiesFactory.GetJobShippingColorPersistenceStrategy().GetCollection());
        }
        public void EditJob(Job job)
        {
            PersistenceStrategiesFactory.GetJobPersistenceStrategy().Edit(job);
        }
        public List<Printer> GetPrinters()
        {
            return new List<Printer>(PersistenceStrategiesFactory.GetPrinterPersistenceStrategy().GetCollection());
        }
        public void EditPrinter(Printer printer)
        {
            PersistenceStrategiesFactory.GetPrinterPersistenceStrategy().Edit(printer);
        }

        public List<User> GetUsers()
        {
            return new List<User>(PersistenceStrategiesFactory.GetUserPersistenceStrategy().GetCollection());
        }
        #endregion


        public void EditJobColors(List<JobColor> jobColors)
        {
            List<Job> jobs = GetJobs();
            foreach (JobColor jobColor in jobColors)
            {
                Job job = jobs.FirstOrDefault(x => x.JOB == jobColor.Job);

                if (string.IsNullOrEmpty(jobColor.Color))
                    job.ShippingLabelColor = "#FFC0CB";
                else
                {
                    job.ShippingLabelColor = jobColor.Color;
                }
                EditJob(job);
            }
        }
        public void EditJobColorsCollection(List<JobColor> jobColors)
        {
            List<Job> jobs = GetJobs();
            foreach (JobColor jobColor in jobColors)
            {
                List<Job> job = new List<Job>(jobs.Where(x => x.Parent == jobColor.Job));
                job.ForEach(x => {
                    x.ShippingLabelColor = jobColor.Color;
                    EditJob(x);
                });
            }
        }
        public void AssignNewColor(JobColor jobColor)
        {
            List<Job> jobs = GetJobs();

            Job job = jobs.FirstOrDefault(x => x.JOB == jobColor.Job);

            job.ShippingLabelColor = jobColor.Color;

            EditJob(job);
        }
        public void ChangePrinterStatus(int line, bool active)
        {
            Printer printer = GetPrinters().FirstOrDefault(x => x.GlazingLine == line);

            printer.Active = active;

            EditPrinter(printer);
        }
        public User ValidateUser(string userName, string password)
        {
            return GetUsers().FirstOrDefault(x => x.Username == userName && x.Password == password);
        }
    }
}
