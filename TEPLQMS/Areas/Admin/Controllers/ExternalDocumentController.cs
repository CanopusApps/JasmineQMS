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
using TEPLQMS.Common;
using TEPLQMS.Controllers;

namespace TEPLQMS.Areas.Admin.Controllers
{
    public class ExternalDocumentController : Controller
    {
        // GET: Admin/ExternalDocument
        [CustomAuthorize(Roles = "QPADM")]
        public ActionResult Index()
        {
            QMSAdmin objAdm = new QMSAdmin();
            ViewBag.Data = objAdm.GetAllExternalDocuments();
            return View();
        }
        [CustomAuthorize(Roles = "ADMIN,QPADM,QMSADMIN")]
        public ActionResult Create()
        {
            QMSAdmin objAdm = new QMSAdmin();
            int seqNo = objAdm.GetExternalDocumentSeqNumber();
            ViewBag.SequenceNumber = seqNo;
            ViewBag.DocumentNumber = "TEPL-EXT-" + seqNo.ToString().PadLeft(4, '0');
            return View();
        }

        public ActionResult CreateBKP()
        {
            return View();
        }

        [CustomAuthorize(Roles = "ADMIN,QPADM,QMSADMIN")]
        public ActionResult Edit()
        {
            ViewBag.ExternalDocumentID = new Guid(Request.QueryString["ID"].ToString());
            return View();
        }
        //public ActionResult GetExternalDocument(int ExternalDocumentID)
        //{
        //    try
        //    {
        //        var list1 = GetExternalDocumentForID(ExternalDocumentID);
        //        return Json(new { success = true, message1 = list1 });
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new { success = false, message1 = "Error" });
        //    }
        //}


        [HttpPost]
        public ActionResult GetSubCategoriesForCategory(string catID)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<ExternalSubCategory> list1 = objAdm.GetExternalSubCategoriesForCategory(new Guid(catID));
                return Json(new { success = true, message1 = list1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetUserDepartment(string loginID)
        {
            try
            {
                string dept = "";
                LoginUser obj = CommonMethods.GetUserADValuesDirectoryEntry(loginID);
                if (obj != null)
                {
                    dept = obj.DepartmentName;
                }
                return Json(new { success = true, message = dept }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetUserDepartmentUsingPrinContext(string loginID)
        {
            try
            {
                string dept = "";
                LoginUser obj = CommonMethods.GetUserADValuesDirectoryEntryUsingPrinContext(loginID);
                if (obj != null)
                {
                    dept = obj.DepartmentName;
                }
                return Json(new { success = true, message = dept }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        [CustomAuthorize(Roles = "ADMIN,QPADM,QMSADMIN")]
        public ActionResult Details()
        {
            try
            {
                ViewBag.ExternalDocumentID = new Guid(Request.QueryString["ID"].ToString());
                return View();
            }
            catch (Exception ex)
            {

            }
            return View();
        }
        [CustomAuthorize(Roles = "ADMIN,QPADM,QMSADMIN")]
        public ActionResult GetExternalDocumentForID(string ExternalDocumentID)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                var data = objAdm.GetExternalDocumentForID(new Guid(ExternalDocumentID));
                return Json(new { success = true, message = data }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "error while getting data" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult SearchUser(string UserName)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<User> listUsers = objAdm.SearchUser(UserName);
                return Json(new { success = true, message = listUsers }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult SaveExternalDocument()
        {
            string response = "";
            ExternalDocument obj = new ExternalDocument();
            byte[] byteArray = null;
            bool docUploaded = false;
            try
            {
                obj.CategoryID = new Guid(Request.Form["Category"].ToString());
                if (Request.Form["SubCategory"].ToString() != "" && Request.Form["SubCategory"].ToString() != "0")
                    obj.SubCategoryID = new Guid(Request.Form["SubCategory"].ToString());
                obj.Title = Request.Form["Title"].ToString();
                obj.DocumentNo = Request.Form["DocumentNo"].ToString();
                if (Request.Form["ExternalDocumentNo"].ToString() != "")
                {
                    obj.ExternalDocumentNo = Request.Form["ExternalDocumentNo"].ToString();
                }
                else
                    obj.ExternalDocumentNo = "";
                if (Request.Form["SequenceNumber"].ToString() != "")
                {
                    obj.SequenceNumber = Convert.ToInt32(Request.Form["SequenceNumber"].ToString());
                }
                else
                    obj.SequenceNumber = 1;
                if (Request.Form["Version"].ToString() != "")
                    obj.Version = Request.Form["Version"].ToString();
                if (Request.Form["VersionDate"].ToString() != "")
                    obj.VersionDate = DateTime.Parse(Request.Form["VersionDate"].ToString());
                obj.Organization = Request.Form["Organization"].ToString();
                if (Request.Form["ResponsibleUser"].ToString() != "")
                    obj.ResponsibleUserID = new Guid(Request.Form["ResponsibleUser"].ToString());
                if (Request.Form["Department"].ToString() != "")
                    obj.Department = Request.Form["Department"].ToString();
                else
                    obj.Department = "";
                obj.FileURL = Request.Form["FileURL"].ToString();
                obj.CreatedBy = new Guid(System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString());
                obj.DocumentName = "";
                obj.FileName = "";
                obj.FilePath = "";
                obj.Active = true;

                //save file to  folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];

                    string flname = "";
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
                        byteArray = fileByteArray;
                        obj.DocumentName = file.FileName;
                        obj.FileName = Request.Form["DocNameNoExt"].ToString() + "_" + obj.DocumentNo + "." + Request.Form["DocNameExt"].ToString();
                        docUploaded = true;
                        if (Request.Form["SubCategory"].ToString() != "" && Request.Form["SubCategory"].ToString() != "0")
                        {
                            obj.FilePath = obj.CategoryID.ToString() + "/" + obj.SubCategoryID.ToString();
                        }
                        else
                        {
                            obj.FilePath = obj.CategoryID.ToString();
                        }
                        //FileName = file.FileName;
                    }
                    // Get the complete folder path and store the file inside it.  
                    //if (flname != "")
                    //{
                    //    //flname = Path.Combine(Server.MapPath("~/assets/img/avatars/"), flname);
                    //    flname = Path.Combine(Server.MapPath("~/Files/"), flname);
                    //    file.SaveAs(flname);
                    //}
                }
                QMSAdmin objAdmin = new QMSAdmin();
                response = objAdmin.AddExternalDocument(obj, byteArray, docUploaded);
            }
            catch (Exception ex)
            {
                response = "Error while saving the External Document";
            }

            return Json(new { success = true, message = response }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateExternalDocument()
        {
            string response = "";
            ExternalDocument obj = new ExternalDocument();
            byte[] byteArray = null;
            bool docUploaded = false;
            try
            {
                obj.ID = new Guid(Request.Form["ExternalDocumentID"].ToString());
                obj.Title = Request.Form["Title"].ToString();
                obj.DocumentNo = Request.Form["DocumentNo"].ToString();
                if (Request.Form["ExternalDocumentNo"].ToString() != "")
                    obj.ExternalDocumentNo = Request.Form["ExternalDocumentNo"].ToString();
                if (Request.Form["Version"].ToString() != "")
                    obj.Version = Request.Form["Version"].ToString();
                if (Request.Form["VersionDate"].ToString() != "")
                    obj.VersionDate = DateTime.Parse(Request.Form["VersionDate"].ToString());
                obj.Organization = Request.Form["Organization"].ToString();
                if (Request.Form["ResponsibleUserID"].ToString() != "")
                    obj.ResponsibleUserID = new Guid(Request.Form["ResponsibleUserID"].ToString());
                obj.Department = Request.Form["Department"].ToString();
                obj.FileURL = Request.Form["FileURL"].ToString();
                obj.ModifiedBy = new Guid(System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString());
                obj.FileName = Request.Form["OLDFileName"].ToString();
                obj.DocumentName = Request.Form["OldDocumentName"].ToString();
                obj.FilePath = Request.Form["FilePath"].ToString();
                //save file to  folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];

                    string flname = "";
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
                        byteArray = fileByteArray;
                        obj.DocumentName = file.FileName;
                        obj.FileName = Request.Form["DocNameNoExt"].ToString() + "_" + obj.DocumentNo + "." + Request.Form["DocNameExt"].ToString();
                        docUploaded = true;
                        if (Request.Form["SubCategory"].ToString() != "" && Request.Form["SubCategory"].ToString() != "0")
                        {
                            obj.FilePath = Request.Form["Category"].ToString() + "/" + Request.Form["SubCategory"].ToString();
                        }
                        else
                        {
                            obj.FilePath = Request.Form["Category"].ToString();
                        }
                    }
                }
                obj.Active = true;
                QMSAdmin objAdmin = new QMSAdmin();
                response = objAdmin.UpdateExternalDocument(obj, byteArray, docUploaded);
            }
            catch (Exception ex)
            {
                response = "Error while saving the External Document";
            }

            return Json(new { success = true, message = response }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteExternalDocument()
        {
            string result = "";

            try
            {
                Guid ID = new Guid(Request.Form["ID"].ToString());
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteExternalDocument(ID);
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
                message = "Document Deleted Successfully."
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DownloadDocument(string filePath, string fileName)
        {
            try
            {
                string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.ExtDocumentFolder, filePath, fileName);
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