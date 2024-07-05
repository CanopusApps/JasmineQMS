
using System;
using System.Web.Mvc;
using TEPL.QMS.Common;

namespace TEPLQMS.Controllers
{
    public class HomeController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();

        [CustomAuthorize()]
        public ActionResult Index()
        {
            try
            {
                LoggerBlock.WriteLog("In Home Controller and in try");
                //int x = 0;
                //int y = 5;
                //int z = y / x; 
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return View();
        }

        public ActionResult Databoxes()
        {
            LoggerBlock.WriteLog("In Home Controller and in Databoxes view");
            return View();
        }

        public ActionResult DataTables()
        {
            LoggerBlock.WriteLog("In Home Controller and in DataTables view");
            return View();
        }

        public ActionResult AllControls()
        {
            return View();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
