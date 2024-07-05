using TEPLQMS.Models.Database;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common;

namespace TEPLQMS.Common
{
    public class LoginManager
    {
         public LoginManager()
        {
            if(HttpContext.Current.Session[QMSConstants.LoggedInUserID] != null)
            {
                AppUser user = (AppUser)HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                ID = user.ID;
                FirstName = user.FirstName;
                LastName = user.LastName;
                Email = user.Email;
                Password = user.Password;
                Mobile = user.Mobile;
                DOB = user.DOB;
                Gender = user.Gender;
                Picture = user.Picture;
                IsActive = user.IsActive;
            }
            else
            {
                FormsAuthentication.SignOut();
                HttpContext.Current.Response.Redirect("~/Login");
                LoggerBlock.WriteLog("Logged In session got expired");                
                throw new Exception("sessionexpired");
            }
        }

        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string Mobile { get; set; }
        public DateTime? DOB { get; set; }
        public string Gender { get; set; }
        public string Picture { get; set; }
        public bool? IsActive { get; set; }
    }
}