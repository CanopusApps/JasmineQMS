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
    public class RequestController : Controller
    {
        public static Guid CurrentStageID = new Guid("25a1f3c9-a06b-4a23-8e3b-328ca0ae1baa");
        // GET: Request
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
                QMSAdmin objQMSAdmin = new QMSAdmin();
                string strRoles = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString();
                if (strRoles.Contains("QADM"))
                    ViewBag.isQMSAdmin = true;
                else ViewBag.isQMSAdmin = false;

                if (strRoles.Contains("QPADM"))
                    ViewBag.isProjectAdmin = true;
                else ViewBag.isProjectAdmin = false;
                string stage = "";
                if(Request.QueryString["stage"] != null)
                {
                    stage = Request.QueryString["stage"].ToString();
                }
                DraftDocument document = null;
                if(stage.ToLower() == "completed")
                {
                    document = obj.GetPublishedDocumentDetailsByID(loggedUsedID, new Guid(strID), false);
                }
                else
                    document = obj.GetDocumentDetailsByID("User", loggedUsedID, new Guid(strID));

                ViewBag.LoggedUsedID = loggedUsedID;
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

        [CustomAuthorize(Roles = "QADM,USER")]
        public ActionResult Details()
        {
            DocumentUpload obj = new DocumentUpload();
            string strID = ""; bool isHistory = false; bool isArchived = false;
            if (Request.QueryString["ID"] != null)
                strID = Request.QueryString["ID"].ToString();
            if (Request.QueryString["IsHistory"] != null)
                isHistory = Convert.ToBoolean(Request.QueryString["IsHistory"].ToString());
            if (Request.QueryString["IsArchived"] != null)
                isArchived = Convert.ToBoolean(Request.QueryString["IsArchived"].ToString());
            ViewBag.isHistory = isHistory;
            ViewBag.isArchived = isArchived;
            if (strID != "")
            {
                Guid loggedUsedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                ViewBag.Data = obj.GetPublishedDocumentDetailsByID(loggedUsedID, new Guid(strID), isHistory);

                string strRoles = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString();
                if (strRoles.Contains("QADM"))
                    ViewBag.isQMSAdmin = true;
                else ViewBag.isQMSAdmin = false;

                DraftDocument document = obj.GetDocumentDetailsByNoForRequest(ViewBag.Data.DocumentNo);
                ViewBag.IsInProgress = document.WFStatus.ToString();
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
        [CustomAuthorize(Roles = "QADM,USER")]
        public ActionResult HistoryDetails()
        {
            DocumentUpload obj = new DocumentUpload();
            string strID = ""; bool isHistory = false;
            if (Request.QueryString["ID"] != null)
                strID = Request.QueryString["ID"].ToString();
            if (Request.QueryString["IsHistory"] != null)
                isHistory = Convert.ToBoolean(Request.QueryString["IsHistory"].ToString());
            ViewBag.isHistory = isHistory;
            if (strID != "")
            {
                Guid loggedUsedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                ViewBag.Data = obj.GetPublishedDocumentDetailsByID(loggedUsedID, new Guid(strID), isHistory);

                string strRoles = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString();
                if (strRoles.Contains("QADM"))
                    ViewBag.isQMSAdmin = true;
                else ViewBag.isQMSAdmin = false;

                DraftDocument document = obj.GetDocumentDetailsByNoForRequest(ViewBag.Data.DocumentNo);
                ViewBag.IsInProgress = document.WFStatus.ToString();
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
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult History()
        {
            DocumentUpload obj = new DocumentUpload();
            string strID = "";
            if (Request.QueryString["ID"] != null)
                strID = Request.QueryString["ID"].ToString();
            if (strID != "")
            {
                Guid loggedUsedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                ViewBag.Data = obj.GetPublishedDocumentHistoryByID(loggedUsedID, new Guid(strID));
            }
            else
            {
                ViewBag.Data = null;
            }
            return View();
        }

        [HttpPost]
        [CustomAuthorize(Roles = "USER")]
        public ActionResult ReSubmitDocument()
        {
            string result = "";
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.DocumentID = new Guid(Request.Form["documentguid"].ToString());
                objDoc.DocumentCategoryCode = Request.Form["doccategorycode"].ToString();
                objDoc.WFExecutionID = new Guid(Request.Form["WFExecutionID"].ToString());
                objDoc.SectionID = new Guid(Request.Form["SectionID"].ToString());
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.Comments = Request.Form["comments"].ToString();
                objDoc.DocumentDescription = Request.Form["DocumentDescription"].ToString();
                objDoc.RevisionReason = Request.Form["RevisionReason"].ToString();
                objDoc.EditableDocumentName = Request.Form["EditableDocumentName"].ToString();
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.UploadedUserName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.CurrentStageID = CurrentStageID;
                objDoc.ProjectTypeID = new Guid(Request.Form["ProjectTypeID"]);
                objDoc.ProjectID = new Guid(Request.Form["ProjectID"]);
                objDoc.DraftVersion = Convert.ToDecimal(Request.Form["draftversion"].ToString());
                objDoc.Action = "Re-submitted";
                bool isDocumentUploaded = false;
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname; //string temFileName = "";

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
                            objDoc.EditableFilePath = Request.Form["filepath"].ToString();
                        }
                        else if (i == 1)
                        {
                            objDoc.ReadableByteArray = fileByteArray;
                            objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.ReadableFilePath = Request.Form["filepath2"].ToString();
                        }
                        isDocumentUploaded = true;
                    }                    
                }

                DocumentUpload bllOBJ = new DocumentUpload();
                bllOBJ.ReSubmitDocument(objDoc, isDocumentUploaded);
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

        [CustomAuthorize(Roles = "QADM")]
        public ActionResult DeleteDocument(string DocumentID)
        {
            try
            {
                Guid UserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objAdm = new QMSAdmin();
                string strMessage = objAdm.DeleteDocument(UserID,new Guid(DocumentID));
                return Json(new { success = true, message = strMessage }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [CustomAuthorize(Roles = "QADM")]
        public ActionResult ArchiveDocument(string DocumentID, string DocumentNo)
        {
            try
            {
                Guid UserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                string UserName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                QMSAdmin objAdm = new QMSAdmin();
                string strMessage = objAdm.ArchiveDocument(UserID, new Guid(DocumentID), DocumentNo, UserName);
                return Json(new { success = true, message = strMessage }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult EditDocument()
        {
            try
            {
                DocumentUpload bllOBJ = new DocumentUpload();
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.UploadedUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.UploadedUserName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                bool isDocumentUploaded = false; string Comments = "";
                if (Request.Form["Comments"].ToString() != "")
                    Comments = Request.Form["Comments"].ToString();
                if (Request.Form["EditableDocumentUploaded"].ToString() == "yes")
                    isDocumentUploaded = true;

                string flname=""; byte[] fileByteArray=null;
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                if (isDocumentUploaded)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];                        
                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            flname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            flname = file.FileName;
                            fileByteArray = new byte[file.ContentLength];
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
                    string EditableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName);
                    string ReadableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName);
                    objDoc.EditableByteArray = bllOBJ.DownloadDocument(EditableURL);
                    objDoc.ReadableByteArray = bllOBJ.DownloadDocument(ReadableURL);
                }
                bllOBJ.ReplaceDocument(objDoc, isDocumentUploaded, Comments);
                return Json(new { success = true, message = "Documents Replaced successfully." }, JsonRequestBehavior.AllowGet);
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}