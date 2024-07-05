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
    public class ProjectUsersController : Controller
    {
        // GET: ProjectUsers
        [CustomAuthorize(Roles = "QADM,QPADM")]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetProjects()
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();

                List<ProjectType> list = objAdm.GetProjectTypesbyRole("ApprovalMatrix", System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString(),
                       (List<Project>)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects]);
                //var list2 = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects];                
                List<Project> list2 = objAdm.GetProjectsbyRole("ProjectUsers",System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserRoles].ToString(),
                       (List<Project>)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserProjects]);
                return Json(new { success = true, message1 = list, message2 = list2 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult SearchProjectUsers(string ProjectTypeID, string ProjectID, string UserName)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                //List<User> listUsers = objAdm.SearchUser(UserName);
                List<User> listUsers = objAdm.SearchProjectUsers(new Guid(ProjectTypeID), new Guid(ProjectID), UserName, false);
                return Json(new { success = true, message = listUsers }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetProjectUsersData(string ProjectTypeID, string ProjectID)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<User> list = objAdm.GetProjectUsers(new Guid(ProjectTypeID), new Guid(ProjectID));
                return Json(new { success = true, message = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult InsertRecord(List<string> arr)
        {
            string result = "";
            try
            {
                if (arr[0].ToString() == "")
                    result = InsertEntity(arr);
                else
                    result = UpdateEntity(arr);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "Issue while adding user";
            }

            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string InsertEntity(List<string> arr)
        {
            string strReturn = string.Empty;
            try
            {
                QMSAdmin objBLL = new QMSAdmin();
                strReturn = objBLL.AddUserToProject(new Guid(arr[1].ToString()), new Guid(arr[2].ToString()), arr[3].ToString().Trim(','), arr[4].ToString().Trim(','));
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw;
            }
            return strReturn;
        }
        [HttpPost]
        public ActionResult DeleteProjectUser(string UserID, string ProjectTypeID, string ProjectID)
        {
            string result = "";
            try
            {
                string LoggedInUserID = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                string strMessage = "";
                if (LoggedInUserID != UserID)
                {
                    QMSAdmin objBLL = new QMSAdmin();

                    strMessage = objBLL.DeleteUserFromProject(new Guid(ProjectTypeID), new Guid(ProjectID), new Guid(UserID));
                }
                else
                {
                    strMessage = "You can not delete your ID.";
                }

                return Json(new { success = true, message = strMessage }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "-1";
            }

            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }
        private string UpdateEntity(List<string> arr)
        {
            try
            {

            }
            catch
            {
                throw;
            }
            return arr[0].ToString();
        }
    }
}