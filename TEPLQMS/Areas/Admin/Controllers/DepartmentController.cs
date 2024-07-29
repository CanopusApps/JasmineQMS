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
    
    public class DepartmentController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetDepartments();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }

        [HttpPost]
        public ActionResult AddDepartment()
        {
            string result = "";
            try
            {
                Departments objDept = new Departments();
                if (Request.Form["DepartmentCode"].ToString() != "")
                    objDept.Code = Request.Form["DepartmentCode"].ToString();
                if (Request.Form["DepartmentName"].ToString() != "")
                    objDept.Title = Request.Form["DepartmentName"].ToString();
                if (Request.Form["DepartmentHOD"].ToString() != "")
                    objDept.HODID = new Guid(Request.Form["DepartmentHOD"].ToString());
                QMSAdmin objAdmin = new QMSAdmin();
                result=objAdmin.AddDepartment(objDept);
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
        public ActionResult UpdateDepartment()
        {
            string result = "";
            try
            {
                Departments objDept = new Departments();
                if (Request.Form["DepartmentID"].ToString() != "")
                    objDept.ID = new Guid(Request.Form["DepartmentID"].ToString());
                if (Request.Form["DepartmentCode"].ToString() != "")
                    objDept.Code = Request.Form["DepartmentCode"].ToString();
                if (Request.Form["DepartmentName"].ToString() != "")
                    objDept.Title = Request.Form["DepartmentName"].ToString();
                if (Request.Form["DepartmentHOD"].ToString() != "")
                    objDept.HODID = new Guid(Request.Form["DepartmentHOD"].ToString());
                if (Request.Form["Active"].ToString().ToLower() == "true")
                    objDept.Active = true;
                else
                    objDept.Active = false;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.UpdateDepartment(objDept);
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

        public ActionResult DeleteDepartment(Guid ID)
        {
            string result = "";

            try
            {
                Departments objDept = new Departments();
                objDept.ID = ID;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.DeleteDepartment(objDept);
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