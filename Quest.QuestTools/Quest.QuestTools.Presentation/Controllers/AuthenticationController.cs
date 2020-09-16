using Quest.Framework.Data.Entities;
using Quest.QuestTools.Business;
using Quest.QuestTools.Presentation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quest.QuestTools.Presentation.Controllers
{
    public class AuthenticationController : Controller
    {

        public User User { get; set; }
        // GET: Authentication
        public FacadeController Controller { get; set; }
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public JsonResult LogIn(UserAuthentication data)
        {
            this.User = Controller.ValidateUser(data.username, data.password);
            Session["User"] = this.User;
            
            return Json(this.User, JsonRequestBehavior.AllowGet);
        }
        public ActionResult LogOut()
        {
            Session["User"] = null;
            if (string.IsNullOrEmpty((string)Session["LastView"])) Session["LastView"] = "Index";
            if (string.IsNullOrEmpty((string)Session["LastController"])) Session["LastController"] = "Home";

            return RedirectToAction((string) Session["LastView"], (string)Session["LastController"], null);
        }
    }
}