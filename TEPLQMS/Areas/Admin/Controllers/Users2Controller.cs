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
    public class Users2Controller : Controller
    {
        // GET: ProjectUsers
        [CustomAuthorize(Roles = "QADM")]
        public ActionResult Index()
        {
            QMSAdmin objAdm = new QMSAdmin();
            ViewBag.Data = objAdm.GetUsers();
            return View();
        }

        [HttpPost]
        public JsonResult SearchADUsersWithName(string Name)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<User> listUsers = CommonMethods.GetUsersFromADUsingDirectorySearch(Name);
                return Json(new { success = true, message = listUsers }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetUserData(string id)
        {
            try
            {
                QMSAdmin objAdm = new QMSAdmin();
                List<User> list = objAdm.GetUserData(new Guid(id));
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
                Guid LoggedInUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objBLL = new QMSAdmin();
                bool IsQMSAdmin = false;
                if (arr[4].ToString().ToLower() == "true")
                    IsQMSAdmin = true;
                strReturn = objBLL.AddUser(arr[1].ToString(), arr[2].ToString(), arr[3].ToString(), IsQMSAdmin, LoggedInUserID);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw;
            }
            return strReturn;
        }
        [HttpPost]
        public ActionResult DeleteUser(string UserID)
        {
            string result = "";
            try
            {
                QMSAdmin objBLL = new QMSAdmin();
                string strMessage = objBLL.DeleteUser(new Guid(UserID));

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
            string strReturn = "";
            try
            {
                Guid LoggedInUserID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                QMSAdmin objBLL = new QMSAdmin();
                bool IsQMSAdmin = false;
                if (arr[4].ToString().ToLower() == "true")
                    IsQMSAdmin = true;

                bool isActive = false;
                if (arr[5].ToString().ToLower() == "true")
                    isActive = true;
                else
                    isActive = false;
                strReturn = objBLL.UpdateUser(new Guid(arr[0].ToString()), arr[1].ToString(), arr[2].ToString(), arr[3].ToString(), IsQMSAdmin, isActive, LoggedInUserID);
            }
            catch
            {
                throw;
            }
            return strReturn;
        }
    }
}