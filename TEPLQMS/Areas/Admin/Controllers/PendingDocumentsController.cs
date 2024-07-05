using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;
using TEPLQMS.Controllers;

namespace TEPLQMS.Areas.Admin.Controllers
{
    public class PendingDocumentsController : Controller
    {
        // GET: Dashboard
        [CustomAuthorize(Roles = "QADM,QPADM")]
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult GetDropDownBoxes()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<Departments> list1 = objAdm.GetDepartments();
                List<DocumentCategory> list2 = objAdm.GetDocumentCategories();


                var list3 = objAdm.GetProjectsbyRole("PendingDocuments", System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString(),
                       (List<Project>)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects]);

                return Json(new { success = true, message1 = list1, message2 = list2, message3 = list3 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult GetExportPendingDocuments(string department, string section, string project, string category, string DocumentDescription,bool IsProjectActive)
        {
            Guid LoggedInUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
            DocumentUpload obj = new DocumentUpload();
            byte[] fileContent = obj.GetExportPendingDocuments(department, section, project, category, DocumentDescription, LoggedInUserID, IsProjectActive);
            string base64 = Convert.ToBase64String(fileContent, 0, fileContent.Length);
            return Content(base64);
        }
        public ActionResult GetDocumentsPending(string department, string section, string project, string category, string DocumentDescription,bool IsProjectActive)
        {
            try
            {
                Guid LoggedInUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                DocumentUpload obj = new DocumentUpload();
                var objDocs = obj.GetDraftDocuments(department, section, project, category, DocumentDescription, LoggedInUserID, IsProjectActive);
                return Json(new { success = true, message = objDocs }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult DownloadDraftDocument(string folderPath, double versionNo, string fileName)
        {
            try
            {
                string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, folderPath, fileName);
                DocumentUpload bllOBJ = new DocumentUpload();
                byte[] fileContent = bllOBJ.DownloadDocument(URL);

                return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}