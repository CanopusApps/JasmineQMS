using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;

namespace TEPLQMS.Controllers
{
    public class DownloadTemplateController : Controller
    {
        // GET: DownloadTemplate
        [CustomAuthorize(Roles = "USER")]
        public ActionResult Index()
        {
            TemplateDocumentBLL objtemp = new TemplateDocumentBLL();
            ViewBag.Data = objtemp.GetTemplateDocuments();

            return View();
        }

        public ActionResult DownloadDocument(string filePath, string fileName)
        {
            try
            {
                string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.TemplateFolder, filePath, fileName);
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