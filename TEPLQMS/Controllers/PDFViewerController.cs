using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;

namespace TEPLQMS.Controllers
{
    public class PDFViewerController : Controller
    {
        public string DocumentURL = string.Empty;
        // GET: PDFViewer
        public ActionResult Index(string docURL)
        {
            DocumentURL = ConfigurationManager.AppSettings["PDFViewerURL"].ToString() + "documentURL=" + TempData["docURL"].ToString() + "&random=" + DateTime.UtcNow.ToString();
            //DocumentURL = ConfigurationManager.AppSettings["PDFViewerURL"].ToString() + "documentURL=" + docURL + "&random=" + DateTime.UtcNow.ToString();
            ViewBag.DocURL = DocumentURL;
            

            return View();
        }

        public ActionResult SetURL(string docURL)
        {
            TempData["docURL"] = docURL;
            return RedirectToAction("Index");
        }

        public static string RemoveQueryStringByKey(string url, string key)
        {
            var uri = new Uri(url);

            // this gets all the query string key value pairs as a collection
            var newQueryString = HttpUtility.ParseQueryString(uri.Query);

            // this removes the key if exists
            newQueryString.Remove(key);

            // this gets the page path from root without QueryString
            string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path);

            return newQueryString.Count > 0
                ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString)
                : pagePathWithoutQueryString;
        }

        [HttpPost]
        public ActionResult GetURL()
        {
            try
            {                

                return Json(new { success = true, message = "http://localhost:61245/Default.htm?documentURL=PublishedDocuments/2f464a59-1d69-4772-9219-b5928defed60/ReadableDocuments/16b0896b-cc50-4851-9d57-c79aa647b3ff/TCPL-PJ0001-EG-ATE-PCY-0001.pdf" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                
                return Json(new { success = true, message = "" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}