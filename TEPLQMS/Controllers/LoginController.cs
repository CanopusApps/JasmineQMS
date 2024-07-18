using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Data.SqlClient;
using System.Text;
using System.Data;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using TEPLQMS.Models.Database;
using TEPLQMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using TEPL.QMS.Common.Models;
using TEPL.QMS.BLL.Component;
using System.Web.Routing;

namespace TEPLQMS.Controllers
{
    public class LoginController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Register()
        {
            return View();
        }

        public ActionResult ForgotPassword()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LoginUsingDirectorySearch(string LoginId, string password)
        {
            try
            {
                //PrincipalContext context = new PrincipalContext(ContextType.Domain);
                //PrincipalContext context = new PrincipalContext(ContextType.Domain, "tataelectronics.co.in");                
                //bool isValid = context.ValidateCredentials(LoginId, password);
                bool isValid = false;
                string loginID = CommonMethods.ValidateUserUsingDirectoryEntry(LoginId, password);
                if(!string.IsNullOrEmpty(loginID))
                {
                    isValid = true;
                }
                var roles = "";
                if (isValid == true)
                {
                    LoginUser obj = new LoginUser();

                    QMSAdmin objAdmin = new QMSAdmin();
                    obj = objAdmin.GetUserDetails(LoginId);
                    if (string.IsNullOrEmpty(obj.LoginID))
                    {
                        if (QMSConstants.SiteName == "OfficeArea")
                        {
                            obj = CommonMethods.GetUserADValuesDirectoryEntry(LoginId);
                            //obj = CommonMethods.GetUserADValuesDirectoryEntryUsingPrinContext(LoginId);
                            objAdmin.InsertUserDetails(obj);
                            obj = objAdmin.GetUserDetails(LoginId);
                        }
                        else
                        {
                            //Access Denied
                            return Json(new { success = true, message = "accessdenied" }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserID, obj.ID);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserDisplayName, obj.DisplayName);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserProjects, obj.Projects);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserRoles, obj.Roles);
                    roles = obj.Roles; //Comma separeted values ex - "USER,ADMIN"

                    FormsAuthentication.SetAuthCookie(QMSConstants.AuthCookieName + "_" + LoginId, false);

                    var authTicket = new FormsAuthenticationTicket(1, QMSConstants.AuthCookieName + "_" + LoginId, DateTime.Now, DateTime.Now.AddMinutes(20), false, roles);
                    string encryptedTicket = FormsAuthentication.Encrypt(authTicket);
                    var authCookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
                    HttpContext.Response.Cookies.Add(authCookie);

                    LoggerBlock.WriteLog("User for the email " + LoginId + " cookie is created.");
                    return Json(new { success = true, message = "sucess" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //User Credentials are in Valid.
                    return Json(new { success = false, message = "fail" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "exception" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult LoginUsingPrincipalContext(string LoginId, string password)
        {
            try
            {
                //PrincipalContext context = new PrincipalContext(ContextType.Domain);
                PrincipalContext context = new PrincipalContext(ContextType.Domain, "tataelectronics.co.in");
                bool isValid = context.ValidateCredentials(LoginId, password);
                var roles = "";
                if (isValid == true)
                {
                    LoginUser obj = new LoginUser();

                    QMSAdmin objAdmin = new QMSAdmin();
                    obj = objAdmin.GetUserDetails(LoginId);
                        if (string.IsNullOrEmpty(obj.LoginID))
                    {
                        if (QMSConstants.SiteName == "OfficeArea")
                        {
                            obj = CommonMethods.GetUserADValuesDirectoryEntry(LoginId);
                            //obj = CommonMethods.GetUserADValuesDirectoryEntryUsingPrinContext(LoginId);
                            objAdmin.InsertUserDetails(obj);
                            obj = objAdmin.GetUserDetails(LoginId);
                        }
                        else
                        {
                            //Access Denied
                            return Json(new { success = true, message = "accessdenied" }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserID, obj.ID);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserDisplayName, obj.DisplayName);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserProjects, obj.Projects);
                    System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserRoles, obj.Roles);
                    roles = obj.Roles; //Comma separeted values ex - "USER,ADMIN"

                    FormsAuthentication.SetAuthCookie(QMSConstants.AuthCookieName + "_" + LoginId, false);

                    var authTicket = new FormsAuthenticationTicket(1, QMSConstants.AuthCookieName + "_" + LoginId, DateTime.Now, DateTime.Now.AddMinutes(20), false, roles);
                    string encryptedTicket = FormsAuthentication.Encrypt(authTicket);
                    var authCookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
                    HttpContext.Response.Cookies.Add(authCookie);

                    LoggerBlock.WriteLog("User for the email " + LoginId + " cookie is created.");
                    return Json(new { success = true, message = "sucess" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //User Credentials are in Valid.
                    return Json(new { success = false, message = "fail" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "exception" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult Logout()
        {
            string landingSite = ConfigurationManager.AppSettings["LandingWebsite"].ToString();
            FormsAuthentication.SignOut();
            HttpContext.Session.Remove(QMSConstants.LoggedInUserID);
            HttpContext.Session.Remove(QMSConstants.LoggedInUserDisplayName);
            //return RedirectToAction("Index", "Login");
            return Redirect(landingSite);
        }

        [HttpPost]
        public ActionResult GetPassword(string mobile, string email)
        {
            string strPassword = "";
            try
            {
                LoggerBlock.WriteLog("User requested for Password. Email: " + email + ". Mobile: " + mobile);
                var sUser = db.AppUsers.Where(a => a.Email.ToLower() == email.ToLower() && a.Mobile == mobile).Select(x => x.Password).First();

                if (sUser != null)
                {
                    strPassword = sUser.ToString();
                    if (strPassword != "")
                    {
                        try
                        {
                            SendEmail(email, strPassword);
                        }
                        catch (Exception ex)
                        {
                            LoggerBlock.WriteTraceLog(ex);
                            return Json(new { success = true, message = "execp" }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    return Json(new { success = true, message = "sucess" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = true, message = "fail" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "fail" }, JsonRequestBehavior.AllowGet);
            }
        }

        private static void SendEmail(string toemail, string password)
        {
            //send mail
            string subject = "Your Password";

            try
            {
                StringBuilder body = new StringBuilder();
                body.Append("<html>");
                body.Append("<body>");
                body.Append("Hello,");
                body.Append("<br/><br/>");
                body.Append("Your password is " + password + ". ");
                body.Append("<br/>");
                body.Append("Please <a href='/Login.aspx'>click here.</a> to login.");
                body.Append("<br/><br/>");
                body.Append("Thanks & Regards,");
                body.Append("<br/>");
                body.Append("Admin Team");

                body.Append("</body>");
                body.Append("</html>");
                MailSend.SendMail(toemail, "", "", subject, "", body.ToString());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}