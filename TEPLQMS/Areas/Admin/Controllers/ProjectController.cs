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
    
    public class ProjectController : Controller
    {
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                ViewBag.Data = objAdm.GetProjects();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
            return View();
        }

        [HttpPost]
        public ActionResult GetProjectTypes()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                var list = objAdm.GetProjectTypes();

                return Json(new { success = true, message1 = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult AddProject(List<string> arr)
        {
            string result = "";
            try
            {
                Project objProj = new Project();
                if (arr[1].ToString() != "")
                    objProj.ProjectTypeID = new Guid(arr[1].ToString());
                if (arr[2].ToString() != "")
                    objProj.WorkflowID = new Guid(arr[2].ToString());
                if (arr[3].ToString() != "")
                    objProj.Code = arr[3].ToString();
                if (arr[4].ToString() != "")
                    objProj.Title = arr[4].ToString();
                if (arr[5].ToString().ToLower() == "true")
                    objProj.ProjectActive = true;
                else
                    objProj.ProjectActive = false;
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.AddProject(objProj);
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
        public ActionResult UpdateProject(List<string> arr)
        {
            string result = "";
            try
            {
                Project objProj = new Project();
                if (arr[0].ToString() != "")
                    objProj.ID = new Guid(arr[0].ToString());                
                if (arr[1].ToString() != "")
                    objProj.Title = arr[1].ToString();
                if (arr[2].ToString().ToLower() == "true")
                    objProj.ProjectActive = true;
                else
                    objProj.ProjectActive = false;
                objProj.ModifiedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objAdmin = new QMSAdmin();
                result = objAdmin.UpdateProject(objProj);
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
    }
}