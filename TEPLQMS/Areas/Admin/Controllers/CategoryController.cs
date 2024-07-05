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
    
    public class CategoryController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetAllDocumentCategories();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }

        [HttpPost]
        public ActionResult AddCategory()
        {
            string result = "";
            try
            {
                DocumentCategory objDocCat = new DocumentCategory();
                if (Request.Form["DocumentCategoryCode"].ToString() != "")
                    objDocCat.Code = Request.Form["DocumentCategoryCode"].ToString();
                if (Request.Form["DocumentCategoryName"].ToString() != "")
                    objDocCat.Title = Request.Form["DocumentCategoryName"].ToString();
                if (Request.Form["DocumentLevel"].ToString() != "")
                    objDocCat.DocumentLevel = Request.Form["DocumentLevel"].ToString();
                //if (Request.Form["FolderName"].ToString() != "")
                //    objDocCat.FolderName = Request.Form["FolderName"].ToString();
                objDocCat.Active = true;
                QMSAdmin objAdmin = new QMSAdmin();
                result=objAdmin.AddDocumentCategory(objDocCat);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
            }
            return Json(new
            {
                success = true,
                message = result
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult UpdateCategory()
        {
            string result = "";
            try
            {
                DocumentCategory objDocCat = new DocumentCategory();
                if (Request.Form["DocumentCategoryID"].ToString() != "")
                    objDocCat.ID = new Guid(Request.Form["DocumentCategoryID"].ToString());
                if (Request.Form["DocumentCategoryName"].ToString() != "")
                    objDocCat.Title = Request.Form["DocumentCategoryName"].ToString();
                if (Request.Form["DocumentLevel"].ToString() != "")
                    objDocCat.DocumentLevel = Request.Form["DocumentLevel"].ToString();
                //if (Request.Form["FolderName"].ToString() != "")
                //    objDocCat.FolderName = Request.Form["FolderName"].ToString();
                if (Request.Form["IsActive"].ToString().ToLower() == "true")
                    objDocCat.Active = true;
                else
                    objDocCat.Active = false;

                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.UpdateDocumentCategory(objDocCat);
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

        public ActionResult DeleteCategory(Guid ID)
        {
            string result = "";
           
            try
            {
                DocumentCategory objDocCat = new DocumentCategory();
                objDocCat.ID = ID;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteDocumentCategory(objDocCat);
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
                message = result
            }, JsonRequestBehavior.AllowGet);
        }
    }
}