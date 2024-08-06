using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPLQMS.Controllers;

namespace TEPLQMS.Areas.Admin.Controllers
{
    public class DocumentNoController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetDocumentNumbers("", "", "", "","");
                //int i =  0,j=1,k=j/i;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }
        public ActionResult GetDocumentNumbers(string department, string section, string project, string category, string function)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                var objDocs = objAdm.GetDocumentNumbers(department, section, project, category, function);
                return Json(new { success = true, message = objDocs }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult UpdateSerialNo(string DocumentID, string SerialNo)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                var result = objAdm.UpdateDocumentNo(new Guid(DocumentID), Convert.ToInt32(SerialNo));
                return Json(new { success = true, message = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "exception" }, JsonRequestBehavior.AllowGet);
            }
        }

    }
}