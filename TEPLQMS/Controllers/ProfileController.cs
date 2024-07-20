using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using TEPLQMS.Models.Database;
using System.Web.Security;
using System.IO;
using TEPLQMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common;

namespace TEPLQMS.Controllers
{
    [CustomAuthorize]
    public class ProfileController : Controller
    {
        // GET: Profile
        public ActionResult Index()
        {
            LoggerBlock.WriteLog("User has come to Proifle Page");
            AppUser usr = null;
            if (Session[QMSConstants.LoggedInUserID] != null)
            {
                usr = (AppUser)Session[QMSConstants.LoggedInUserID];
            }
            ViewBag.cUserID = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.ID.ToString() });
            ViewBag.cUserName = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.FirstName.ToString() });
            ViewBag.cEmail = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.Email.ToString() });
            ViewBag.cMobile = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.Mobile.ToString() });
            string img = "";
            if (!string.IsNullOrEmpty(usr.Picture))
                img = usr.Picture;
            ViewBag.cPic = JsonConvert.SerializeObject(new ViewEnable { Enable = img });
            return View();
        }


        [HttpPost]
        public ActionResult ChangeProfile(HttpPostedFileBase fil)
        {
            string result = ""; AppUser usr = null;
            LoggerBlock.WriteLog("User trying to change the Proifle details.");
            try
            {
                if (Session[QMSConstants.LoggedInUserID] != null)
                {
                    usr = (AppUser)Session[QMSConstants.LoggedInUserID];
                }
                int id = Convert.ToInt32(Request.Form["id"].ToString());
                string name = Request.Form["name"].ToString();
                string mail = Request.Form["mail"].ToString();
                string mob = Request.Form["mob"].ToString();
                string picname = Request.Form["picname"].ToString();

                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string fname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        fname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        fname = file.FileName;
                    }
                    // Get the complete folder path and store the file inside it.  
                    if (fname != "")
                    {
                        fname = Path.Combine(Server.MapPath("~/assets/img/avatars/"), fname);
                        file.SaveAs(fname);
                    }
                }

                result = UpdateProfile(id,mail,name,mob, usr.FirstName, picname);
                LoggerBlock.WriteLog("User has changed the Proifle details.");
                usr.ID = id;
                usr.Email = mail;
                usr.Password = usr.Password;
                usr.FirstName = name;
                usr.Mobile = mob;
                usr.Picture = picname;
                usr.IsActive = usr.IsActive;

                HttpContext.Session[QMSConstants.LoggedInUserID] = usr;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string UpdateProfile(int id, string email, string name, string mob, string modname,string fileName)
        {
            string result = "";
            try
            {                
                using (DataBaseEntities ctx = new DataBaseEntities())
                {
                    var query = (from q in ctx.AppUsers
                                 where q.ID == id
                                 select q).First();
                    query.Email = email;
                    query.FirstName = name;
                    query.Mobile = mob;
                    query.Picture = fileName;
                    query.ModifiedBy = modname;
                    query.ModifiedDate = DateTime.Now;
                    ctx.SaveChanges();
                    result = "changed";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
            base.Dispose(disposing);
        }
    }
}