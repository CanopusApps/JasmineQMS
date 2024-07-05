using TEPLQMS.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using TEPL.QMS.Common.Constants;
using MvcSiteMapProvider.Security;

namespace TEPLQMS.Controllers
{
    public class CustomAuthorize : AuthorizeAttribute
    {
        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            try
            {
                base.OnAuthorization(filterContext);
                if (!filterContext.HttpContext.User.Identity.IsAuthenticated)
                {
                    filterContext.Result = new RedirectResult("~/Login");
                    return;
                }

                if (HttpContext.Current.Session[QMSConstants.LoggedInUserID] == null)
                {
                    filterContext.Result = new RedirectResult("~/Login");
                    return;
                }

                //Form Authentication
                var authCookie = HttpContext.Current.Request.Cookies[FormsAuthentication.FormsCookieName];
                if (authCookie != null)
                {
                    FormsAuthenticationTicket authTicket = FormsAuthentication.Decrypt(authCookie.Value);

                    if (authTicket == null || authTicket.Expired || (!authTicket.Name.Contains(QMSConstants.AuthCookieName)))
                    {
                        filterContext.Result = new RedirectResult("~/Login");
                        return;
                    }
                }

                if (filterContext.Result is HttpUnauthorizedResult)
                {
                    filterContext.Result = new RedirectResult("~/Error/NotAuthorized");
                    return;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}