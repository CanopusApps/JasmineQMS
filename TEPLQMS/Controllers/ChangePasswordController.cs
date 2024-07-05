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
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common;
using TEPLQMS.Common;

namespace TEPLQMS.Controllers
{
    public class ChangePasswordController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();
    
        // GET: ChangePassword
        public ActionResult Index()
        {
            LoggerBlock.WriteLog("User has come to Change password Link");            
            ViewBag.cUserID = JsonConvert.SerializeObject(new ViewEnable { Enable = (new LoginManager()).ID.ToString() });
            ViewBag.cUserName = JsonConvert.SerializeObject(new ViewEnable { Enable = (new LoginManager()).FirstName.ToString() });
            ViewBag.cPassword = JsonConvert.SerializeObject(new ViewEnable { Enable = (new LoginManager()).Password.ToString() });

            return View();
        }


        [HttpPost]
        public ActionResult ChangePassowrd(List<string> arr)
        {
            string result = ""; 
            try
            {
                LoggerBlock.WriteLog("User trying to change the password for the ID: " + arr[0].ToString());
                result = UpdatePassowrdNew(Convert.ToInt32(arr[0]), arr[1].ToString(), arr[2].ToString());

                //Kill the Session & Cookie
                FormsAuthentication.SignOut();
                HttpContext.Session.Remove(QMSConstants.LoggedInUserID);

                LoggerBlock.WriteLog("User has changed the password for the ID: " + arr[0].ToString());
            }
            catch (Exception ex)
            {
                result = "failed";
               LoggerBlock.WriteTraceLog(ex);
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string UpdatePassowrdNew(int id, string psw, string name)
        {
            string result = "";
            try
            {
                using (DataBaseEntities ctx = new DataBaseEntities())
                {
                    var query = (from q in ctx.AppUsers
                                 where q.ID == id
                                 select q).First();
                    query.Password = psw;
                    query.ModifiedBy = name;
                    query.ModifiedDate = DateTime.Now;
                    ctx.SaveChanges();
                    result = "changed";
                }

                //This will also work but all mandatory fields has to mention while creating obj
                //var _person = new SiteUser() { ID = id, Password = psw };

                //using (var newContext = new DataBaseEntities())
                //{
                //    newContext.UserNAs.Attach(_person);
                //    newContext.Entry(_person).Property(X => X.Password).IsModified = true;
                //    newContext.SaveChanges();
                //    result = "changed";
                //}
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return result;
        }

        #region OldCode
        private string UpdatePassowrd(int id, string psw, string name)
        {
            string result = "";
            try
            {
                AppUser user = new AppUser();
                user.Password = psw;
                user.ID = id;
                user.IsActive = true;
                user.ModifiedBy = name;
                user.ModifiedDate = DateTime.Now;
                user.Email = "";

                if (ModelState.IsValid)
                {
                    db.AppUsers.Attach(user);
                    db.Entry(user).Property(x => x.Email).IsModified = false;
                    db.Entry(user).Property(x => x.FirstName).IsModified = false;
                    db.Entry(user).Property(x => x.Mobile).IsModified = false;
                    db.Entry(user).Property(x => x.CreatedBy).IsModified = false;
                    db.Entry(user).Property(x => x.CreatedDate).IsModified = false;

                    db.Entry(user).Property(x => x.ModifiedBy).IsModified = true;
                    db.Entry(user).Property(x => x.ModifiedDate).IsModified = true;
                    db.Entry(user).Property(x => x.Password).IsModified = true;
                    db.Entry(user).Property(x => x.IsActive).IsModified = true;

                    db.SaveChanges();

                    HttpContext.Session.Remove(QMSConstants.LoggedInUserID);
                    FormsAuthentication.SignOut();
                    result = "changed";
                }                
            }
            catch (Exception ex)
            {
                result = "failed";
            }
            finally
            {
            }
            return result;
        }

        private string UpdatePassowrdOLD(int id, string psw)
        {
            string result = "";
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("[UpdatePassword]");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@Password", psw);

                string strConnString = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                        //Remove the Session
                        HttpContext.Session.Remove("cUser");
                        result = "changed";
                    }
                }
            }
            catch (Exception ex)
            {
                result = "failed";
            }
            finally
            {
                cmd.Dispose();
            }
            return result;
        }
        #endregion

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