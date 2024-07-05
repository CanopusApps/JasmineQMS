using System;
using System.Collections.Generic;
using System.IO;
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
    public class TemplateDocumentController : Controller
    {
        // GET: Admin/TemplateDocument
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                TemplateDocumentBLL objTempBLL = new TemplateDocumentBLL();
                ViewBag.Data = objTempBLL.GetTemplateDocuments();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }
        public ActionResult AddTemplate()
        {
            string result = "";
            try
            {
                TemplateDocument objTemp = new TemplateDocument();
                if (Request.Form["TemplateCode"].ToString() != "")
                    objTemp.Code = Request.Form["TemplateCode"].ToString();
                if (Request.Form["TemplateName"].ToString() != "")
                    objTemp.Name = Request.Form["TemplateName"].ToString();
                if (Request.Form["TemplateLevel"].ToString() != "")
                    objTemp.Level = Request.Form["TemplateLevel"].ToString();
                objTemp.CreatedID= (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                    }
                    else
                    {
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        objTemp.byteArray = fileByteArray;
                        objTemp.FileName = file.FileName;
                        objTemp.FilePath = "";
                    }
                }
                TemplateDocumentBLL objTempBLL = new TemplateDocumentBLL();
                objTempBLL.AddTemplate(objTemp);
                result = "sucess";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new
            {
                success = true,
                message = result
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteTemplate()
        {
            string result = "";

            try
            {
                Guid ID = new Guid(Request.Form["ID"].ToString());
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteTemplate(ID);
                // result = "sucess";

            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                throw ex;
            }
            return Json(new
            {
                success = true,
                message = "Template Deleted Successfully."
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateTemplate()
        {
            string result = ""; bool isFileUploaded = false;
            try
            {
                TemplateDocument objTemp = new TemplateDocument();
                if (Request.Form["ID"].ToString() != "")
                    objTemp.ID = new Guid(Request.Form["ID"].ToString()); 
                if (Request.Form["TemplateCode"].ToString() != "")
                    objTemp.Code = Request.Form["TemplateCode"].ToString();
                if (Request.Form["TemplateName"].ToString() != "")
                    objTemp.Name = Request.Form["TemplateName"].ToString();
                if (Request.Form["TemplateLevel"].ToString() != "")
                    objTemp.Level = Request.Form["TemplateLevel"].ToString();
                if (Request.Form["ExtFileName"].ToString() != "")
                    objTemp.FileName = Request.Form["ExtFileName"].ToString();
                objTemp.ModifiedBy = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                    }
                    else
                    {
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        objTemp.byteArray = fileByteArray;
                        objTemp.FileName = file.FileName;
                        objTemp.FilePath = "";
                        isFileUploaded = true;
                    }
                }
                //QMSAdmin objQMSAdmin = new QMSAdmin();
                //objQMSAdmin.UpdateTemplate(objTemp);
                TemplateDocumentBLL objTempBLL = new TemplateDocumentBLL();
                objTempBLL.UpdateTemplate(objTemp, isFileUploaded);
                result = "sucess";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new
            {
                success = true,
                message = result
            }, JsonRequestBehavior.AllowGet);
        }
    }
}