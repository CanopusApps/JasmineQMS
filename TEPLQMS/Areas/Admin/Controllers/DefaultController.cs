using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPLQMS.Controllers;

namespace TEPLQMS.Areas.Admin.Controllers
{
    public class DefaultController : Controller
    {
        // GET: Admin/Default
        [CustomAuthorize()]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetDepartments();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }
    }
}