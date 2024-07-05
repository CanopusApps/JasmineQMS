using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using System.Configuration;
using Newtonsoft.Json;
using TEPL.QMS.Common.Models;
using TEPL.QMS.Common.Constants;
using System.IO;
using TEPL.QMS.Common.Model;
using TEPL.QMS.Common;
using System.Text;
using System.Net.Mail;
using System.Security.Cryptography;

namespace TEPLQMS.Controllers
{
    public class DocumentInitiateController : Controller
    {
        // GET: DocumentInitiate
        public static Guid CurrentStageID = new Guid("f5d83a73-c4a3-48aa-b3d7-568301eea27b");

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
        public ActionResult GetDropDownBoxes()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<Departments> list1 = objAdm.GetDepartments();
                List<DocumentCategory> list2 = objAdm.GetDocumentCategories();

                var list3 = objAdm.GetProjectsbyRole("DocumentInitiate", System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString(),
                       (List<Project>)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects]);

                return Json(new { success = true, message1 = list1, message2 = list2, message3 = list3 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetSectionsForDepartment(string deptID)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<Sections> list1 = objAdm.GetSectionsForDept(new Guid(deptID));
                return Json(new { success = true, message1 = list1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult GenerateDocumentNumber()
        {
            try
            {
                Object[] ArrayOfObjects = new Object[3];
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.WorkflowID = QMSConstants.WorkflowID;
                objDoc.CurrentStageID = CurrentStageID;
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                DocumentUpload bllOBJ = new DocumentUpload();
                ArrayOfObjects = bllOBJ.GenerateDocumentNo(objDoc);
                return Json(new { success = true, message1 = ArrayOfObjects }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult GenerateDocumentNumber1(string projCode, string deptCode, string secCode, string catCode)
        {
            try
            {
                Object[] ArrayOfObjects = new Object[3];
                //DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                string compCode = QMSConstants.CompanyCode;
                DraftDocument objDoc = new DraftDocument();

                objDoc.CompanyCode = compCode;
                objDoc.DepartmentCode = deptCode;
                objDoc.SectionCode = secCode;
                objDoc.ProjectCode = projCode;
                objDoc.DocumentCategoryCode = catCode;
                objDoc.WorkflowID = QMSConstants.WorkflowID;
                objDoc.CurrentStageID = CurrentStageID;
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                DocumentUpload bllOBJ = new DocumentUpload();
                ArrayOfObjects = bllOBJ.GenerateDocumentNo(objDoc);
                return Json(new { success = true, message1 = ArrayOfObjects }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult SubmitDocument()
        {
            string result = "";
            try
            {
                StringBuilder strMessage = new StringBuilder();
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);

                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                strMessage.AppendLine("UploadedUserID - " + objDoc.UploadedUserID.ToString() + ". ");
                objDoc.UploadedUserName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                strMessage.AppendLine("DocumentID - " + objDoc.DocumentID.ToString() + ". ");
                strMessage.AppendLine("WFExecutionID - " + objDoc.WFExecutionID.ToString() + ". ");
                objDoc.RevisionReason = "";
                objDoc.CurrentStageID = CurrentStageID;
                objDoc.DraftVersion = 0.001m;
                objDoc.Action = "Submitted";
                //DraftDocument objDoc1 =(DraftDocument)BindModels.GetObject(Request.Form);
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
                            objDoc.EditableFilePath = CommonMethods.CombineUrl(objDoc.ProjectID.ToString(), QMSConstants.EditableFolder, objDoc.DepartmentID.ToString(), objDoc.SectionID.ToString(), objDoc.DocumentCategoryID.ToString());
                        }
                        else if (i == 1)
                        {
                            objDoc.ReadableByteArray = fileByteArray;
                            objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.ReadableFilePath = CommonMethods.CombineUrl(objDoc.ProjectID.ToString(), QMSConstants.ReadableFolder, objDoc.DepartmentID.ToString(), objDoc.SectionID.ToString(), objDoc.DocumentCategoryID.ToString());
                        }
                    }
                }

                DocumentUpload bllOBJ = new DocumentUpload();
                bllOBJ.SubmitDocument(objDoc);

                result = "sucess";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DownloadTemplate(string folderPath, string fileName)
        {
            string result = "";
            //string filename = "TEPL-COMMON-CE-XX-PO-0006.docx";
            //URL = @"file//D://Storage/QMS/PublishDocuments/CEO OFFICE/CEO OFFICE/POLICY/TEPL-COMMON-CE-XX-PO-0006.docx";
            try
            {
                //string URL = siteURL + "/" + docLib + "/" + folderPath + "/" + fileName;
                string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.ReadableFolder, folderPath, fileName);
                DocumentUpload bllOBJ = new DocumentUpload();
                byte[] fileContent = bllOBJ.DownloadDocument(URL);

                return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);

                //return Json(new { success = true, message = "Test" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }
    }
}