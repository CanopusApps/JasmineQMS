using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace TEPLQMS.Controllers
{
    public class ErrorController : Controller
    {
        // GET: Error
        public ActionResult Index()
        {
            return View();
        }

        // GET: Error
        public ActionResult NotAuthorized()
        {
            return View();
        }
        public ActionResult PageNotFoundError()
        {
            return View();
        }
        public ActionResult UnauthorizedError()
        {
            return View();
        }
        public ActionResult InternalServerError()
        {
            return View();
        }
        public ActionResult GenericError()
        {
            return View();
        }
    }
}