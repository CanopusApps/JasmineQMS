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
    
    public class SectionController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetAllSections();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }

        [HttpPost]
        public ActionResult GetDepartments()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<Departments> list1 = objAdm.GetActiveDepartments();
                return Json(new { success = true, message1 = list1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult AddSection()
        {
            string result = "";
            try
            {
                Sections entObject = new Sections();
                if (Request.Form["SectionCode"].ToString() != "")
                    entObject.Code = Request.Form["SectionCode"].ToString();
                if (Request.Form["SectionName"].ToString() != "")
                    entObject.Title = Request.Form["SectionName"].ToString();
                if (Request.Form["DepartmentID"].ToString() != "")
                    entObject.DepartmentID = new Guid(Request.Form["DepartmentID"].ToString());
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.AddSection(entObject);
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

        [HttpPost]
        public ActionResult UpdateSection()
        {
            string result = "";
            try
            {
                Sections entObject = new Sections();
                if (Request.Form["SectionID"].ToString() != "")
                    entObject.ID = new Guid(Request.Form["SectionID"].ToString());
                if (Request.Form["SectionCode"].ToString() != "")
                    entObject.Code = Request.Form["SectionCode"].ToString();
                if (Request.Form["SectionName"].ToString() != "")
                    entObject.Title = Request.Form["SectionName"].ToString();
                if (Request.Form["Active"].ToString().ToLower() == "true")
                    entObject.Active = true;
                else
                    entObject.Active = false;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.UpdateSection(entObject);
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

        public ActionResult DeleteSection(Guid ID)
        {
            string result = "";

            try
            {
                Sections entObject = new Sections();
                entObject.ID = ID;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteSection(entObject);
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