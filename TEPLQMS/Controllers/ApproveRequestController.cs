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
    public class ApproveRequestController : Controller
    {
        // GET: ApproveRequest
        [CustomAuthorize(Roles = "USER")]
        public ActionResult Index()
        {
            DocumentUpload obj = new DocumentUpload();
            string strID = "";
            if (Request.QueryString["ID"] != null)
                strID = Request.QueryString["ID"].ToString();
            if (strID != "")
            {
                Guid loggedUsedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                DraftDocument document = obj.GetDocumentDetailsByID("Approver", loggedUsedID, new Guid(strID));
                ViewBag.Data = document;
                if (document.DraftVersion > 0)
                {
                    ViewBag.DraftVersion = Math.Ceiling(document.DraftVersion);
                }
                else
                {
                    ViewBag.DraftVersion = 0;
                }
            }
            else
            {
                ViewBag.Data = null;
            }
            ViewBag.FileTypes = ConfigurationManager.AppSettings["FileTypes"].ToString();
            ViewBag.FormsFileTypes = ConfigurationManager.AppSettings["FormsFileTypes"].ToString();
            ViewBag.ReadableFileTypes = ConfigurationManager.AppSettings["ReadableFileTypes"].ToString();
            ViewBag.AllowedFileSize = ConfigurationManager.AppSettings["AllowedFileSize"].ToString();
            ViewBag.ViewerURL = ConfigurationManager.AppSettings["ViewerURL"].ToString();
            return View();
        }

        [HttpPost]
        public ActionResult RejectRequest(string docNumber, string docGUID, string comments, string CurrentStageID, string CurrentStage, string exeID, string uplodUserID, string DocumentDescription,
            string DocumentName, string RevisionReason, string DraftVersion,string EditVersion, string OriginalVersion)
        {
            try
            {
                string actBy = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                DraftDocument objDoc = new DraftDocument();
                objDoc.DocumentID = new Guid(docGUID);
                objDoc.DocumentNo = docNumber;
                //objDoc.Comments = comments;
                objDoc.CurrentStageID = new Guid(CurrentStageID);
                objDoc.CurrentStage = CurrentStage;
                objDoc.WFExecutionID = new Guid(exeID);
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.Action = "Rejected";
                objDoc.ActionComments = comments;
                objDoc.UploadedUserID = new Guid(uplodUserID);// System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                objDoc.EditableDocumentName = DocumentName;
                objDoc.DocumentDescription = DocumentDescription;
                if (DraftVersion != "")
                    objDoc.DraftVersion = Convert.ToDecimal(DraftVersion);
                if (EditVersion != "")
                    objDoc.EditVersion = Convert.ToDecimal(EditVersion);
                if (OriginalVersion != "")
                    objDoc.OriginalVersion = Convert.ToDecimal(OriginalVersion);
                if (RevisionReason != null)
                    objDoc.RevisionReason = RevisionReason;
                else
                    objDoc.RevisionReason = "";
                DocumentUpload bllOBJ = new DocumentUpload();
                bllOBJ.RejectDocument(objDoc, false);
                return Json(new { success = true, message1 = "sucess" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "fail" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult ApproveDocument()
        {
            string result = "";
            bool status = true;
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.Action = "Approved";
                bool isDocumentUploaded = false;
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname; string temFileName = "";

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
                            objDoc.EditableFilePath = Request.Form["EditableFilePath"].ToString();
                        }
                        else if (i == 1)
                        {
                            objDoc.ReadableByteArray = fileByteArray;
                            objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.ReadableFilePath = Request.Form["ReadableFilePath"].ToString();
                        }
                        isDocumentUploaded = true;
                    }
                }
                //MultipleApprovers = "{\"DocumentReviewers\":\"E86E0C22-6419-4CAD-941A-BC69E99030E4,EB9523FD-B564-426C-A2B9-53B16522073B\",\"DocumentApprovers\":\"A6E1B539-A1E1-4275-B3E5-D64F967A07AB,5E9D535F-2AC4-430F-9624-ECEE3F789512\"}";
                DocumentUpload bllOBJ = new DocumentUpload();
                bool IsMultiApproversChanged = Convert.ToBoolean(Request.Form["IsMultiApproversChanged"].ToString());
                result = bllOBJ.ApproveDocument(objDoc, isDocumentUploaded, IsMultiApproversChanged);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = status, message = result }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult PublishDocument()
        {
            string result = "";
            try
            {
                DocumentUpload bllOBJ = new DocumentUpload();
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.Action = "Published";
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                bool isDocumentUploaded = false;
                if (Request.Form["EditableDocumentUploaded"].ToString() == "yes")
                    isDocumentUploaded = true;
             
                if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                {
                    objDoc.ReadableDocumentName = Request.Form["EditableDocumentName"].ToString();
                }                    
                else
                {
                    objDoc.ReadableDocumentName = Request.Form["ReadableDocumentName"].ToString();
                }
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                if(isDocumentUploaded)
                {
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
                                if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                                {
                                    objDoc.ReadableByteArray = fileByteArray;
                                    objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                                }
                            }
                            else
                            {
                                objDoc.ReadableByteArray = fileByteArray;
                            }
                        }
                    }
                }
                else
                {
                    string EditableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName);
                    string ReadableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName);
                    if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                    {
                        ReadableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.ReadableDocumentName);
                    }
                    objDoc.EditableByteArray = bllOBJ.DownloadDocument(EditableURL);
                    objDoc.ReadableByteArray = bllOBJ.DownloadDocument(ReadableURL);
                }
                bllOBJ.DocumentPublish(objDoc, isDocumentUploaded);
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

        //public ActionResult DownloadDraftDocument(string folderPath, double versionNo, string fileName)
        //{
        //    //string filename = "TEPL-COMMON-CE-XX-PO-0006.docx";
        //    //URL = @"file//D://Storage/QMS/DraftDocuments/CEO OFFICE/CEO OFFICE/POLICY/TEPL-COMMON-CE-XX-PO-0006.docx";
        //    try
        //    {
        //        //string URL = siteURL + "/" + docLib + "/" + folderPath + "/" + fileName;
        //        string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, folderPath, fileName);
        //        DocumentUpload bllOBJ = new DocumentUpload();
        //        byte[] fileContent = bllOBJ.DownloadDocument(URL);
        //        string contentType = MimeMapping.GetMimeMapping(fileName);

        //        var cd = new System.Net.Mime.ContentDisposition
        //        {
        //            FileName = fileName,
        //            Inline = true,
        //        };

        //        Response.AppendHeader("Content-Disposition", cd.ToString());

        //        return File(fileContent, contentType);
        //        //return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        LoggerBlock.WriteTraceLog(ex);
        //        return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
        //    }
        //}

        //public ActionResult DownloadReadableDocument(string folderPath, double versionNo, string fileName)
        //{
        //    try
        //    {
        //        string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, folderPath, fileName);
        //        DocumentUpload bllOBJ = new DocumentUpload();
        //        byte[] fileContent = bllOBJ.DownloadDocument(URL);

        //        return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        LoggerBlock.WriteTraceLog(ex);
        //        return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
        //    }
        //}


        [HttpPost]
        public ActionResult ConvertoPDF_Test()
        {
            string response = "";
            try
            {

            }
            catch (Exception ex)
            {
                response = "error";
            }
            return Json(new { success = true, message = response }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ConvertToPDF()
        {
            string result = "";
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.Action = "Published";
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname; string temFileName = "";

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        flname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        flname = file.FileName;
                        temFileName = flname;
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        objDoc.ReadableByteArray = fileByteArray;
                        string extension = Path.GetExtension(file.FileName);
                        objDoc.ReadableDocumentName = objDoc.DocumentNo + extension;
                        objDoc.ReadableFilePath = CommonMethods.CombineUrl(objDoc.ProjectName, objDoc.DepartmentName, objDoc.SectionName, objDoc.DocumentCategoryName);
                        objDoc.EditableFilePath = CommonMethods.CombineUrl(objDoc.ProjectName, objDoc.DepartmentName, objDoc.SectionName, objDoc.DocumentCategoryName);
                        FileConvert.ConvertToPDF(file.FileName, fileByteArray);
                    }
                }
                //Logic to convert document to PDF and send the path of link
                result = "/ApproveRequest/DownloadDocument?folderPath=NNB102%2FAdmin%2FADMIN%20COMMON%2FPolicies&versionNo=0.001&fileName=TEPL-NNB102-AD-AC-PO-0007.docx";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }
    }
}