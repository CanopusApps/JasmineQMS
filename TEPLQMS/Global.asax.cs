using TEPLQMS.Common;
using TEPLQMS.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Security;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common.Models;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using Microsoft.Practices.EnterpriseLibrary.Logging;

namespace TEPLQMS
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            Logger.SetLogWriter(new LogWriterFactory().Create());
        }

        protected void Application_PostAuthenticateRequest(Object sender, EventArgs e)
        {
            var authCookie = HttpContext.Current.Request.Cookies[FormsAuthentication.FormsCookieName];
            if (authCookie != null)
            {
                FormsAuthenticationTicket authTicket = FormsAuthentication.Decrypt(authCookie.Value);
                if (authTicket != null && !authTicket.Expired && authTicket.Name.Contains(QMSConstants.AuthCookieName))
                {
                    var roles = authTicket.UserData.Split(',');
                    HttpContext.Current.User = new System.Security.Principal.GenericPrincipal(new FormsIdentity(authTicket), roles);
                }
            }
        }

        //protected void Session_Start(object sender, EventArgs e)
        //{
        //    string[] newroles = null;
        //    try
        //    {
        //        string[] r = { "QMSApprover" };
        //        HttpContext.Current.User = new GenericPrincipal(HttpContext.Current.User.Identity, r);
        //        User obj = new User();
        //        obj = CommonMethods.GetUserADValues(@System.Web.HttpContext.Current.Request.LogonUserIdentity.Name);

        //        obj.CreatedBy = obj.LoginID;


        //        QMSAdmin objAdmin = new QMSAdmin();
        //        obj = objAdmin.GetUserDetails(obj);
        //        if (!string.IsNullOrEmpty(obj.LoginID))
        //        {
        //            System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserID, obj.LoginID);
        //            System.Web.HttpContext.Current.Session.Add(QMSConstants.LoggedInUserDisplayName, obj.DisplayName);
        //            //var gIdentity = new GenericIdentity(vcnID);
        //            //if (HttpContext.Current.Session[Constants.SessionKey_CurrentUser] == null)
        //            //{
        //            //    LoggerBlock.Write("In Session");
        //            //    CommonMethods cm = new CommonMethods();
        //            //    string roleName = cm.GetUserDetails(vcnID);
        //            //    LoggerBlock.Write("Role - " + roleName);
        //            //    if (roleName == "")
        //            //    {
        //            //        throw new Exception("There is a problem with the VCN ID.");
        //            //    }
        //            //    else
        //            //    {
        //            //        newroles = new string[1];
        //            //        newroles[0] = roleName;
        //            //        System.Web.HttpContext.Current.User = new GenericPrincipal(gIdentity, newroles);
        //            //    }
        //            //}
        //        }
        //        else
        //        {
        //            LoggerBlock.WriteLog("Login ID " + @System.Web.HttpContext.Current.Request.LogonUserIdentity.Name + " does not have access to this portal");

        //            var controller = new ErrorController
        //            {
        //                //ViewData = { Model = new HandleErrorInfo(new HttpException(401, "Unauthorized access"), "AccessDenied", "Index") }
        //            };
        //            var routeData = new RouteData();
        //            routeData.Values["controller"] = "AccessDenied";
        //            routeData.Values["action"] = "Index";
        //            ((IController)controller).Execute(new RequestContext(new HttpContextWrapper(System.Web.HttpContext.Current), routeData));
        //            Response.End();
        //            HttpContext.Current.Session.Clear();
        //            HttpContext.Current.Session.Abandon();
        //            //HttpContext.Current.Response.Cookies.Add(new HttpCookie("ASP.NET_SessionId", ""));
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //Log.Error("There is an error while getting the user details from DB", ex);
        //        throw ex;
        //    }
        //    finally
        //    {
        //        newroles = null;
        //    }
        //}

        /// <summary>
        /// Handle the application error redirecting into the error page.
        /// </summary>
        protected void Application_Error(object sender, EventArgs e)
        {
            HttpApplication application = (HttpApplication)sender;

            // Get information about the exception detected for the current request
            // At this stage we know it is not authorization related exception
            var exception = Server.GetLastError();

            // Clear all the errors collected so far
            application.Context.ClearError();

            // Clear the current response and write a new one. For any other error handling
            // (other than authorization) the status code is alway 500 (InternalServerError)
            HttpResponse response = application.Response;
            response.Clear();
            response.StatusCode = (int)HttpStatusCode.InternalServerError;

            var controller = new ErrorController
            {
                //ViewData = { Model = new HandleErrorInfo(exception, controllerName, actionName) }
            };

            var routeData = new RouteData();
            routeData.Values["controller"] = "Error";
            routeData.Values["action"] = "Index";
            ((IController)controller).Execute(new RequestContext(new HttpContextWrapper(application.Context), routeData));

            Response.End();
        }
    }
}
