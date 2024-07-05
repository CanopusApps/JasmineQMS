using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;

namespace TEPLQMS.Areas.Admin.Controllers
{
    public class ExternalSubCategoryController : Controller
    {
        // GET: Admin/ExternalSubCategory
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetExternalSubCategories();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return View();
        }

        [HttpPost]
        public ActionResult GetActiveCategories()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<ExternalCategory> list1 = objAdm.GetActiveExternalCategories();
                return Json(new { success = true, message1 = list1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult AddSubCategory()
        {
            string result = "";
            try
            {
                ExternalSubCategory obj = new ExternalSubCategory();
                if (Request.Form["ExtCategoryID"].ToString() != "")
                    obj.ExtCategoryID = new Guid(Request.Form["ExtCategoryID"].ToString());
                if (Request.Form["CategoryName"].ToString() != "")
                    obj.Title = Request.Form["CategoryName"].ToString();
                obj.CreatedBy = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.AddExternalSubCategory(obj);
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
        public ActionResult UpdateSubCategory()
        {
            string result = "";
            try
            {
                ExternalSubCategory obj = new ExternalSubCategory();
                if (Request.Form["EntityID"].ToString() != "")
                    obj.ID = new Guid(Request.Form["EntityID"].ToString());
                if (Request.Form["CategoryName"].ToString() != "")
                    obj.Title = Request.Form["CategoryName"].ToString();
                if (Request.Form["Active"].ToString().ToLower() == "true")
                    obj.Active = true;
                else
                    obj.Active = false;
                obj.ModifiedBy = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.UpdateExternalSubCategory(obj);
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

        public ActionResult DeleteSubCategory(Guid ID)
        {
            string result = "";

            try
            {
                ExternalSubCategory obj = new ExternalSubCategory();
                obj.ID = ID;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteExternalSubCategory(obj);
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
    }
}