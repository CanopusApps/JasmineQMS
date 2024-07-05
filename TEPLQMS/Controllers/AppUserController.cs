using TEPLQMS.Models.Database;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.Common;
using TEPLQMS.Common;

namespace TEPLQMS.Controllers
{
    public class AppUserController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();
     
        [CustomAuthorize()]
        public ActionResult Index()
        {
            db = new Models.Database.DataBaseEntities();
            var data = db.AppUsers;//.Where(a => a.IsActive == true);

            ViewBag.Data = data;
            return View();
        }


        [HttpPost]
        public ActionResult SaveEntity(List<string> arr, List<string> arr2)
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
                        var res = db.AppUsers.Where(a => a.ID == id).First();
                        mesg = res.FirstName + "#" + res.Email + "#" + res.Mobile;
                        return Json(new { success = true, message = mesg }, JsonRequestBehavior.AllowGet);
                    }
                }
                if (arr[0].ToString() == "")
                    result = InsertEntity(arr, arr2);
                else
                    result = UpdateEntity(arr, arr2);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "-1";
            }

            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        private string InsertEntity(List<string> arr, List<string> arr2)
        {
            int id = 0;
            AppUser entity = null;
            try
            {
                //AppUser user = (AppUser)Session[MyConstants.LoggedInUser];
                string loggedName = (new LoginManager()).FirstName;
                entity = new AppUser();
                entity.FirstName = arr[1].ToString();
                entity.LastName = arr[4].ToString();
                entity.Email = arr[2].ToString();
                entity.Password = ConfigurationManager.AppSettings["DefaultPassword"].ToString();                
                entity.Mobile = arr[5].ToString();
                if (arr[3] != "")
                    entity.DOB = DateTime.Parse(arr[3].ToString());
                if (arr[6].ToString() == "yes")
                    entity.Gender = "Male";
                else
                    entity.Gender = "Female";
                if (arr[8].ToString() == "checked")
                    entity.IsActive = true;
                else
                    entity.IsActive = false;
                entity.CreatedBy = loggedName;
                entity.CreatedDate = DateTime.Now;
                if (ModelState.IsValid)
                {
                    db.AppUsers.Add(entity);
                    db.SaveChanges();
                    id = entity.ID;
                }

                if (id != 0 && arr2 != null)//Add roles for the user
                {
                    for (int i = 0; i < arr2.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(arr2[i]))
                        {
                            AppUserRole childEntity = new AppUserRole();
                            childEntity.AppUserID = id;
                            childEntity.AppRoleID = Convert.ToInt32(arr2[i].ToString());
                            childEntity.CreatedBy = loggedName;
                            childEntity.CreatedDate = DateTime.Now;
                            childEntity.IsActive = true;
                            if (ModelState.IsValid)
                            {
                                db.AppUserRoles.Add(childEntity);
                            }
                        }
                    }
                    db.SaveChanges();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                entity = null;
            }
            return id.ToString();
        }

        private string UpdateEntity(List<string> arr, List<string> arr2)
        {
            AppUser entity = null;
            int id = Convert.ToInt32(arr[0].ToString());
            try
            {
                entity = new AppUser();
                string loggedName = (new LoginManager()).FirstName;
                entity.ID = id;
                entity.FirstName = arr[1].ToString();
                entity.LastName = arr[4].ToString();
                entity.Email = arr[2].ToString();
                entity.Mobile = arr[5].ToString();
                entity.Password = "";
                if (arr[3] != "")
                    entity.DOB = DateTime.Parse(arr[3].ToString());
                if (arr[6].ToString() == "yes")
                    entity.Gender = "Male";
                else
                    entity.Gender = "Female";
                if (arr[8].ToString() == "checked")
                    entity.IsActive = true;
                else
                    entity.IsActive = false;
                entity.ModifiedBy = loggedName;
                entity.ModifiedDate = DateTime.Now;                

                if (ModelState.IsValid)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    db.Entry(entity).Property(x => x.CreatedBy).IsModified = false;
                    db.Entry(entity).Property(x => x.CreatedDate).IsModified = false;
                    db.Entry(entity).Property(x => x.Password).IsModified = false;
                    db.SaveChanges();
                }
                //delete existisng record if any
                db.AppUserRoles.RemoveRange(db.AppUserRoles.Where(c => c.AppUserID == id));

                if (id != 0 && arr2 != null)//Add roles for the user
                {
                    for (int i = 0; i < arr2.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(arr2[i]))
                        {
                            AppUserRole childEntity = new AppUserRole();
                            childEntity.AppUserID = id;
                            childEntity.AppRoleID = Convert.ToInt32(arr2[i].ToString());
                            childEntity.CreatedBy = loggedName;
                            childEntity.CreatedDate = DateTime.Now;
                            childEntity.IsActive = true;
                            if (ModelState.IsValid)
                            {
                                db.AppUserRoles.Add(childEntity);
                            }
                        }
                    }
                    db.SaveChanges();
                }
            }
            catch
            {
                throw;
            }
            return arr[0].ToString();
        }

        private bool IsItemExists(string email, string ID)
        {
            bool IsExists = false;
            try
            {
                var entityID = db.AppUsers.Where(a => a.Email.ToLower().Trim() == email.ToLower()).Select(a => a.ID).First();
                if (!string.IsNullOrEmpty(entityID.ToString()))
                {
                    if (ID.ToString() != entityID.ToString())
                        IsExists = true;
                }
            }
            catch
            {                
                IsExists = false;                
            }
            return IsExists;
        }

        [HttpPost]
        public ActionResult DeleteEntity(int id)
        {
            try
            {
                string loggedName = (new LoginManager()).FirstName;
                var query = (from q in db.AppUsers
                             where q.ID == id
                             select q).First();
                query.IsActive = false;
                query.ModifiedBy = loggedName;
                query.ModifiedDate = DateTime.Now;
                db.SaveChanges();
            }
            catch (Exception ex)
            {
               LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "Failed" }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { success = true, message = "Deleted" }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetAllRoles()
        {
            try
            {
                var list = db.AppRoles;
                return Json(new { success = true, message = list }, JsonRequestBehavior.AllowGet);
            }
            catch(Exception ex)
            {
               LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetDataForMasterID(int id)
        {
            try
            {
                var listRoles = db.AppUserRoles.Where(a => a.AppUserID == id).Select(a => a.AppRoleID);
                var list = db.AppUsers.Where(i => i.ID == id).First();
                //var result = list + "#" + listRoles;
                return Json(new { message = list, success = listRoles }, JsonRequestBehavior.AllowGet);
            }
            catch(Exception ex)
            {
               LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
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