using TEPLQMS.Common;
using TEPLQMS.Models.Database;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common;

namespace TEPLQMS.Controllers
{
    public class TestUserController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();
        [CustomAuthorize(Roles = "Administrator")]
        public ActionResult Index()
        {
            db = new Models.Database.DataBaseEntities();
            var data = db.TestUsers.Where(a => a.IsActive == true);

            ViewBag.Data = data;

            return View();
        }

        private DataTable GetEmployeeDetails()
        {
            DataTable dt = null;
            try
            {
                var data = db.TestUsers.Where(a => a.IsActive == true);

                
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            finally
            {

            }
            return dt;
        }

        [HttpPost]
        public ActionResult SaveEntity(List<string> arr)
        {
            string result = "";
            try
            {
                if (IsItemExists(arr[2].ToString().ToLower().Trim(), arr[0].ToString()))
                {
                    if (arr[0].ToString() == "")
                        return Json(new { success = true, message = "0" }, JsonRequestBehavior.AllowGet);
                    else
                    {
                        string mesg = "";
                        int id = Convert.ToInt32(arr[0].ToString());
                        var res = db.TestUsers.Where(a => a.ID == id).First();
                        mesg = res.FullName + "#" + res.Email + "#" + res.Mobile;
                        return Json(new { success = true, message = mesg }, JsonRequestBehavior.AllowGet);
                    }
                }
                if (arr[0].ToString() == "")
                    result = InsertEntity(arr);
                else
                    result = UpdateEntity(arr);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "Error while saving User.";
            }

            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string InsertEntity(List<string> arr)
        {
            int id = 0;
            TestUser entity = null;
            try
            {
                AppUser user = (AppUser)Session[QMSConstants.LoggedInUserID];
                entity = new TestUser();
                entity.Email = arr[2].ToString();
                entity.Password = ConfigurationManager.AppSettings["DefaultPassword"].ToString();
                entity.FullName = arr[1].ToString();
                entity.Mobile = arr[3].ToString();
                entity.IsActive = true;
                entity.CreatedBy = user.FirstName;
                entity.CreatedDate = DateTime.Now;
                if (ModelState.IsValid)
                {
                    db.TestUsers.Add(entity);
                    db.SaveChanges();
                    id = entity.ID;
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            finally
            {
                entity = null;
            }
            return id.ToString();
        }

        private string UpdateEntity(List<string> arr)
        {
            string result = "User updated successfully.";
            TestUser entity = null;
            try
            {
                entity = new TestUser();
                AppUser user = (AppUser)Session[QMSConstants.LoggedInUserID];
                entity = new TestUser();
                entity.Email = arr[2].ToString();
                entity.FullName = arr[1].ToString();
                entity.Mobile = arr[3].ToString();
                entity.IsActive = true;
                entity.ModifiedBy = user.FirstName;
                entity.ModifiedDate = DateTime.Now;
                entity.ID = Convert.ToInt32(arr[0].ToString());
                entity.Password = "";

                if (ModelState.IsValid)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    db.Entry(entity).Property(x => x.CreatedBy).IsModified = false;
                    db.Entry(entity).Property(x => x.CreatedDate).IsModified = false;
                    db.Entry(entity).Property(x => x.Password).IsModified = false;
                    db.SaveChanges();                    
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            finally
            {
            }
            return result;
        }

        private bool IsItemExists(string email, string ID)
        {
            bool IsExists = false;
            try
            {
                var entityID = db.TestUsers.Where(a => a.Email.ToLower().Trim() == email.ToLower()).Select(a => a.ID).First();
                if (!string.IsNullOrEmpty(entityID.ToString()))
                {
                    if (ID.ToString() != entityID.ToString())
                        IsExists = true;
                }
            }
            catch (Exception ex)
            {
                IsExists = false;
                LoggerBlock.WriteTraceLog(ex);
            }
            return IsExists;
        }

        [HttpPost]
        public ActionResult DeleteEntity(int id)
        {
            try
            {
                AppUser user = (AppUser)Session[QMSConstants.LoggedInUserID];
                var query = (from q in db.TestUsers
                             where q.ID == id
                             select q).First();
                query.IsActive = false;
                query.ModifiedBy = user.FirstName;
                query.ModifiedDate = DateTime.Now;
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "Failed" }, JsonRequestBehavior.AllowGet);
            }
            finally
            {

            }
            return Json(new { success = true, message = "Deleted" }, JsonRequestBehavior.AllowGet);
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