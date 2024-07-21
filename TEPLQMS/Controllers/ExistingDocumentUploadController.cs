using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;

namespace TEPLQMS.Controllers
{
    public class ExistingDocumentUploadController : Controller
    {
        public static Guid CurrentStageID = new Guid("f5d83a73-c4a3-48aa-b3d7-568301eea27b");
        // GET: ExistingDocumentUpload
        [CustomAuthorize(Roles = "USER")]
        public ActionResult Index()
        {
            ViewBag.FileTypes = ConfigurationManager.AppSettings["FileTypes"].ToString();
            ViewBag.FormsFileTypes = ConfigurationManager.AppSettings["FormsFileTypes"].ToString();
            ViewBag.ReadableFileTypes = ConfigurationManager.AppSettings["ReadableFileTypes"].ToString();
            ViewBag.AllowedFileSize = ConfigurationManager.AppSettings["AllowedFileSize"].ToString();

            return View();
        }

        [HttpPost]
        public ActionResult GetDocumentDetails(string documentNumber)
        {
            try
            {
                DocumentUpload bllOBJ = new DocumentUpload();
                string userID = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                //Project[] list3 =(Project[])System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects];
                List<Project> list3 = (List<Project>)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects];
                if (documentNumber.Split('-').Length == 1)
                {
                    return Json(new { success = true, message = "invalid" }, JsonRequestBehavior.AllowGet);
                }
                string ProjectCode = documentNumber.Split('-')[1];
                bool blAccess = false;
                foreach(Project pt in list3)
                {
                    if(pt.Code==ProjectCode)
                    {
                        blAccess = true;
                    }
                }
                if(ProjectCode == "QM")
                {
                    blAccess = true;
                }
                if (blAccess)
                {
                    Object[] ArrayOfObjects = new Object[3];
                    ArrayOfObjects = bllOBJ.GetDocumentDetailsByNo(documentNumber);
                    return Json(new { success = true, message = ArrayOfObjects }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = true, message = "noaccess" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult SubmitExistingUpload()
        {
            string result = "";
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                //DraftDocument objDoc1 = CommonMethods.GetObject(Request.Form, objDoc);

                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        flname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        flname = file.FileName;
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        if (i == 0)
                        {
                            objDoc.EditableByteArray = fileByteArray;
                            objDoc.EditableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            //objDoc.EditableFilePath = CommonMethods.CombineUrl(objDoc.ProjectID.ToString(), QMSConstants.EditableFolder, objDoc.DepartmentID.ToString(), objDoc.SectionID.ToString(), objDoc.DocumentCategoryID.ToString());
                            objDoc.EditableFilePath = CommonMethods.CombineUrl(objDoc.ProjectID.ToString(), QMSConstants.EditableFolder);
                        }
                        else if (i == 1)
                        {
                            objDoc.ReadableByteArray = fileByteArray;
                            objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.ReadableFilePath = CommonMethods.CombineUrl(objDoc.ProjectID.ToString(), QMSConstants.ReadableFolder);
                        }
                    }                    
                }
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.UploadedUserName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.CurrentStageID = CurrentStageID;
                objDoc.Action = "Submitted";

                DocumentUpload bllOBJ = new DocumentUpload();
                bllOBJ.SubmitExistingUpload(objDoc);
                result = "sucess";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed"; return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }
    }
}