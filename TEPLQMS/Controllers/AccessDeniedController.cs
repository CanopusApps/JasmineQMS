using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace TEPLQMS.Controllers
{
    public class AccessDeniedController : Controller
    {
        // GET: AccessDenied
        public ActionResult Index()
        {
            ViewBag.SupportEmail = ConfigurationManager.AppSettings["EmailFrom"].ToString();
            return View();
        }
    }
}