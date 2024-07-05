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
    public class Profile2Controller : Controller
    {
        // GET: Profile2
        public ActionResult Index()
        {
            LoggerBlock.WriteLog("User has come to Proifle Page");
            //AppUser usr = null;
            //if (Session[MyConstants.LoggedInUser] != null)
            //{
            //    usr = (AppUser)Session[MyConstants.LoggedInUser];
            //}
            //ViewBag.cUserID = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.ID.ToString() });
            //ViewBag.cUserName = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.FirstName.ToString() });
            //ViewBag.cEmail = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.Email.ToString() });
            //ViewBag.cMobile = JsonConvert.SerializeObject(new ViewEnable { Enable = usr.Mobile.ToString() });
            //string img = "";
            //if (!string.IsNullOrEmpty(usr.Picture))
            //    img = usr.Picture;
            //ViewBag.cPic = JsonConvert.SerializeObject(new ViewEnable { Enable = img });
            return View();
        }


        [HttpPost]
        public ActionResult ChangeProfile(HttpPostedFileBase fil)
        {
            string result = ""; AppUser usr = null;string mail="";
            LoggerBlock.WriteLog("User trying to change the Proifle details.");
            try
            {
                if (Session[QMSConstants.LoggedInUserID] != null)
                {
                    usr = (AppUser)Session[QMSConstants.LoggedInUserID];
                }
                int id = Convert.ToInt32(Request.Form["id"].ToString());
                string fname = Request.Form["fname"].ToString();
                string lname = Request.Form["lname"].ToString();
                mail = Request.Form["mail"].ToString();
                string mob = Request.Form["mob"].ToString();
                string dob = Request.Form["dob"].ToString();
                string gender = Request.Form["gender"].ToString();                
                string picname = Request.Form["picname"].ToString();

                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        flname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        flname = file.FileName;
                    }
                    // Get the complete folder path and store the file inside it.  
                    if (flname != "")
                    {
                        flname = Path.Combine(Server.MapPath("~/assets/img/avatars/"), flname);
                        file.SaveAs(flname);
                    }
                }
                LoggerBlock.WriteLog("User Saved the image and trying to save details in DB. Email: " + mail);

                result = UpdateProfile(id, mail, fname,lname,gender, mob,dob, usr.FirstName, picname);

                LoggerBlock.WriteLog("User Saved the image and details in DB. Email: " + mail);

                usr.ID = id;
                usr.Email = mail;
                usr.Password = usr.Password;
                usr.FirstName = fname;
                usr.LastName = lname;
                usr.Mobile = mob;
                if (dob != "")
                    usr.DOB = DateTime.Parse(dob);
                usr.Gender = gender;
                usr.Picture = picname;
                usr.IsActive = usr.IsActive;

                HttpContext.Session[QMSConstants.LoggedInUserID] = usr;
            }
            catch (Exception ex)
            {
                result = "failed";
                LoggerBlock.WriteTraceLog(ex);
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string UpdateProfile(int id, string email, string fname, string lname,string gender, 
                                    string mob, string dob, string modname, string fileName)
        {
            string result = "";
            try
            {
                using (DataBaseEntities ctx = new DataBaseEntities())
                {
                    var query = (from q in ctx.AppUsers
                                 where q.ID == id
                                 select q).First();
                    query.FirstName = fname;
                    query.LastName = lname;
                    query.Mobile = mob;
                    if (dob != "")
                        query.DOB = DateTime.Parse(dob) ;
                    query.Gender = gender;
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