using Newtonsoft.Json.Schema;
using Quest.Framework.Data.Entities;
using Quest.Framework.Data.Persistence.OLEDB;
using Quest.QuestTools.Business;
using Quest.QuestTools.Presentation.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quest.QuestTools.Presentation.Controllers
{
    public class ShippingController : AuthenticationController
    {
        // GET: Shipping 
        
        public ShippingController()
        {
            this.Controller = new FacadeController(
                new OLEDBPersistenceStrategiesFactory(
                    ConfigurationManager.ConnectionStrings["QuestDB"].ConnectionString
                    )
                );

            
        }
        public ActionResult ShippingLabelColor()
        {
            this.User = Session["User"] == null ? null : (User)Session["User"];
            ViewBag.Title = "Shipping Label Colors";
            Session["LastView"] = "ShippingLabelColor";
            Session["LastController"] = "Shipping";
            return View("ShippingLabelColor", GetShippingLabelColorViewModel());
        }
        public ActionResult Printers()
        {
            this.User = Session["User"] == null ? null : (User)Session["User"];

            Session["LastView"] = "Printers";
            Session["LastController"] = "Shipping";

            return View(new PrintersViewModel()
            {
                Printers = Controller.GetPrinters(),
                User = this.User
            });
        }
        [HttpPost]
        public JsonResult SaveJobColorChanges(List<JobColorChangeViewModel> jobColorChangesList)
        {
            if (jobColorChangesList.Count(x => x.IsModified) > 0)
            {
                Controller.EditJobColors(new List<JobColor>(jobColorChangesList.Where(w => w.IsModified).Select(x => new JobColor()
                {
                    Job = x.Job,
                    Color = x.ShippingLabelColor
                })));
            }
            return Json(GetShippingLabelColorViewModel().JobShippingColors, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult GetJobs()
        {
            List<Job> jobs = new List<Job>(
                Controller.GetJobs().Where(x => 
                    string.IsNullOrEmpty(x.ShippingLabelColor)
                ).OrderBy(y => y.JOB));
            return Json(jobs, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult AssignNewColor(JobColor data)
        {
            Controller.AssignNewColor(data);

            return Json(GetShippingLabelColorViewModel().JobShippingColors, JsonRequestBehavior.AllowGet);
        }
        public JsonResult ActivatePrinter(Printer printerStatus)
        {
            try
            {
                Controller.ChangePrinterStatus(printerStatus.GlazingLine, printerStatus.Active);
                return Json(string.Format("{0} - {1}", printerStatus.GlazingLine, printerStatus.Active), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(ex, JsonRequestBehavior.DenyGet);
            }
        }
        
        private ShippingLabelColorViewModel GetShippingLabelColorViewModel()
        {
            return new ShippingLabelColorViewModel()
            {
                JobShippingColors = Controller.GetJobShippingColors(),
                Colors = new List<ColorCatalogViewModel>(
                    Controller.GetShippingColors().Select(
                        x =>
                        new ColorCatalogViewModel()
                        {
                            ColorHexRGB = x.ColorHexRGB,
                            ColorName = x.ColorName
                        }
                        )),
                Jobs = new List<Job>(Controller.GetJobs()),
                User = this.User
            };
        }
    }
}