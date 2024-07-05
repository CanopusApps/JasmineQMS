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
    public class ExternalCategoryController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetExternalCategories();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return View();
        }

        [HttpPost]
        public ActionResult AddCategory()
        {
            string result = "";
            try
            {
                ExternalCategory obj = new ExternalCategory();
                if (Request.Form["CategoryName"].ToString() != "")
                    obj.Title = Request.Form["CategoryName"].ToString();
                obj.CreatedBy = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.AddExternalCategory(obj);
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
                ExternalCategory obj = new ExternalCategory();
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
                result = objAdmin.UpdateExternalCategory(obj);
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

        public ActionResult DeleteCategory(Guid ID)
        {
            string result = "";

            try
            {
                ExternalCategory obj = new ExternalCategory();
                obj.ID = ID;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteExternalCategory(obj);
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