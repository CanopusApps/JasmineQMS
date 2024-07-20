using PdfSharp.Drawing.Layout;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TEPL.QMS.BLL.Component;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;


namespace TEPLQMS.Controllers
{
    public class ApproveRequestController : Controller
    {
        // GET: ApproveRequest
        [CustomAuthorize(Roles = "USER")]
        public ActionResult Index()
        {
            DocumentUpload obj = new DocumentUpload();
            string strID = "";
            if (Request.QueryString["ID"] != null)
                strID = Request.QueryString["ID"].ToString();
            if (strID != "")
            {
                Guid loggedUsedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                DraftDocument document = obj.GetDocumentDetailsByID("Approver", loggedUsedID, new Guid(strID));
                ViewBag.Data = document;
                if (document.DraftVersion > 0)
                {
                    ViewBag.DraftVersion = Math.Ceiling(document.DraftVersion);
                }
                else
                {
                    ViewBag.DraftVersion = 0;
                }
            }
            else
            {
                ViewBag.Data = null;
            }
            ViewBag.FileTypes = ConfigurationManager.AppSettings["FileTypes"].ToString();
            ViewBag.FormsFileTypes = ConfigurationManager.AppSettings["FormsFileTypes"].ToString();
            ViewBag.ReadableFileTypes = ConfigurationManager.AppSettings["ReadableFileTypes"].ToString();
            ViewBag.AllowedFileSize = ConfigurationManager.AppSettings["AllowedFileSize"].ToString();
            ViewBag.ViewerURL = ConfigurationManager.AppSettings["ViewerURL"].ToString();
            return View();
        }

        public ActionResult ViewDocument()
        {
            
            return View();
        }

        [HttpPost]
        public ActionResult RejectRequest(string docNumber, string docGUID, string comments, string CurrentStageID, string CurrentStage, string exeID, string uplodUserID, string DocumentDescription,
            string DocumentName, string RevisionReason, string DraftVersion,string EditVersion, string OriginalVersion)
        {
            try
            {
                string actBy = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                DraftDocument objDoc = new DraftDocument();
                objDoc.DocumentID = new Guid(docGUID);
                objDoc.DocumentNo = docNumber;
                //objDoc.Comments = comments;
                objDoc.CurrentStageID = new Guid(CurrentStageID);
                objDoc.CurrentStage = CurrentStage;
                objDoc.WFExecutionID = new Guid(exeID);
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.Action = "Rejected";
                objDoc.ActionComments = comments;
                objDoc.UploadedUserID = new Guid(uplodUserID);// System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                objDoc.EditableDocumentName = DocumentName;
                objDoc.DocumentDescription = DocumentDescription;
                if (DraftVersion != "")
                    objDoc.DraftVersion = Convert.ToDecimal(DraftVersion);
                if (EditVersion != "")
                    objDoc.EditVersion = Convert.ToDecimal(EditVersion);
                if (OriginalVersion != "")
                    objDoc.OriginalVersion = Convert.ToDecimal(OriginalVersion);
                if (RevisionReason != null)
                    objDoc.RevisionReason = RevisionReason;
                else
                    objDoc.RevisionReason = "";
                DocumentUpload bllOBJ = new DocumentUpload();
                bllOBJ.RejectDocument(objDoc, false);
                return Json(new { success = true, message1 = "sucess" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                return Json(new { success = false, message = "fail" }, JsonRequestBehavior.AllowGet);
            }
        }

        

        [HttpPost]
        public ActionResult ApproveDocument()
        {
            string result = "";
            bool status = true;
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                objDoc.Action = "Approved";
                bool isDocumentUploaded = false;
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname; //string temFileName = "";

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        flname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        flname = file.FileName;
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        if (i == 0)
                        {
                            objDoc.EditableByteArray = fileByteArray;
                            objDoc.EditableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.EditableFilePath = Request.Form["EditableFilePath"].ToString();
                        }
                        else if (i == 1)
                        {
                            objDoc.ReadableByteArray = fileByteArray;
                            objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                            objDoc.ReadableFilePath = Request.Form["ReadableFilePath"].ToString();
                        }
                        isDocumentUploaded = true;
                    }
                }
                //MultipleApprovers = "{\"DocumentReviewers\":\"E86E0C22-6419-4CAD-941A-BC69E99030E4,EB9523FD-B564-426C-A2B9-53B16522073B\",\"DocumentApprovers\":\"A6E1B539-A1E1-4275-B3E5-D64F967A07AB,5E9D535F-2AC4-430F-9624-ECEE3F789512\"}";
                DocumentUpload bllOBJ = new DocumentUpload();
                bool IsMultiApproversChanged = Convert.ToBoolean(Request.Form["IsMultiApproversChanged"].ToString());


                if(objDoc.CurrentStage == "Document Approver")
                {
                    string filePath = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName);
                    AddWatermarkonPDF(filePath);
                }

                result = bllOBJ.ApproveDocument(objDoc, isDocumentUploaded, IsMultiApproversChanged);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = status, message = result }, JsonRequestBehavior.AllowGet);
        }

        public void AddWatermarkonPDF(string ipFilename)
        {
            try
            {
                //Watermark text
                string cntrWatermark = "TCPL CONFIDENTIAL";
                string rgtTopWatermark = "MP CONFIDENTIAL";
                string rgtBtmWatermark = "TCPL Confidential" + System.Environment.NewLine +
                    "No Copy/Reproduction allowed" + System.Environment.NewLine + System.Environment.NewLine +
                    "CONTROLLED COPY" + System.Environment.NewLine + "ISSUED BY DMS" +
                    System.Environment.NewLine + System.Environment.NewLine + DateTime.Now.ToString("dd-MM-yyyy") +
                    System.Environment.NewLine + "TCPL-DCC";

                //Create a new PDF document
                PdfDocument document = new PdfDocument();

                //string ipFilename = Path.GetFullPath("Data/ToTestPPT1.pdf");
                //string FilePath = Server.MapPath("ToTestPPT1.pdf");
                //string ipFilename = "D:\\Sample.pdf";

                //Open PDF document
                document = PdfReader.Open(ipFilename);

                // Set version to PDF 1.4 (Acrobat 5) because we use transparency.
                if (document.Version < 14)
                    document.Version = 14;

                // Create a font
                XFont cntrFont = new XFont("Arial", 40, XFontStyleEx.Bold);
                XFont rgtFont = new XFont("Arial", 12, XFontStyleEx.Bold);

                for (int idx = 0; idx < document.Pages.Count; idx++)
                {
                    var page = document.Pages[idx];

                    //-- Page Center watermark
                    // Get an XGraphics object for drawing beneath the existing content.
                    var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

                    // Get the size (in points) of the text.
                    var cntrTxtSize = gfx.MeasureString(cntrWatermark, cntrFont);

                    // Define a rotation transformation at the center of the page.
                    gfx.TranslateTransform(page.Width / 2, page.Height / 2);
                    gfx.RotateTransform(-Math.Atan(page.Height / page.Width) * 180 / Math.PI);
                    gfx.TranslateTransform(-page.Width / 2, -page.Height / 2);

                    // Create a string format.
                    var cntrFormat = new XStringFormat();
                    cntrFormat.Alignment = XStringAlignment.Near;
                    cntrFormat.LineAlignment = XLineAlignment.Near;

                    // Create a colored brush.
                    XBrush cntrBrush = new XSolidBrush(XColor.FromArgb(128, 69, 69, 69)); //light black
                    XBrush rgtBrush = new XSolidBrush(XColor.FromArgb(128, 0, 110, 255)); //blue 

                    // Format text and add rectangle border
                    XTextFormatter cntrTf = new XTextFormatter(gfx);
                    XRect cntrRect = new XRect((page.Width - cntrTxtSize.Width) / 2, (page.Height - cntrTxtSize.Height) / 2, 450, 20); //rgt pos, vertical pos of textbox, width, height
                    cntrTf.Alignment = XParagraphAlignment.Center;
                    XPen cntrPen = new XPen(XColors.Black, 1);
                    gfx.DrawRectangle(cntrPen, (page.Width - cntrTxtSize.Width) / 2, (page.Height - cntrTxtSize.Height) / 2, 450, 40);
                    cntrTf.DrawString(cntrWatermark, cntrFont, cntrBrush, cntrRect, cntrFormat);
                    gfx.Dispose();

                    //-- Page Right top watermark
                    // Get an XGraphics object for drawing beneath the existing content.
                    var rgtTopGfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

                    // Get the size (in points) of the text.
                    var rgtTopTxtSize = rgtTopGfx.MeasureString(rgtTopWatermark, rgtFont);

                    // Format text and add rectangle border at right top corner
                    XTextFormatter rgtTopTf = new XTextFormatter(rgtTopGfx);
                    //XRect rgtTopRect = new XRect((page.Width - rgtTopTxtSize.Width) - 33, (rgtTopTxtSize.Height) - 10, 200, 15);//rgt pos, vertical pos of textbox, width, height
                    XRect rgtTopRect = new XRect((page.Width - rgtTopTxtSize.Width) - 35, (rgtTopTxtSize.Height) - 9, rgtTopTxtSize.Width + 62, rgtTopTxtSize.Height);//rgt pos, vertical pos of textbox, width, height
                    rgtTopTf.Alignment = XParagraphAlignment.Center;
                    XPen rgtTopPen = new XPen(XColors.Black, 1);
                    rgtTopGfx.DrawRectangle(rgtTopPen, (page.Width - rgtTopTxtSize.Width) - 5, (rgtTopTxtSize.Height) - 10, rgtTopTxtSize.Width + 2, rgtTopTxtSize.Height + 2);
                    rgtTopTf.DrawString(rgtTopWatermark, rgtFont, rgtBrush, rgtTopRect, XStringFormats.TopLeft);
                    rgtTopGfx.Dispose();

                    ////without rectangle border
                    //rgtTopGfx.DrawString(rgtTopWatermark, rgtFont, rgtBrush,
                    //    new XPoint((page.Width - rgtTopTxtSize.Width) - 5, (rgtTopTxtSize.Height) + 2)); //new XPoint((page.Width - 200), 30));
                    //rgtTopGfx.Dispose();

                    //-- Page Right bottom watermark
                    // Get an XGraphics object for drawing beneath the existing content.
                    var rgtBtmGfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

                    // Format text and add rectangle border at right bottom corner
                    XTextFormatter rgtBtmTf = new XTextFormatter(rgtBtmGfx);
                    XRect rect = new XRect((page.Width - 260), (page.Height - 120), 318, 1020);//rgt pos, vertical pos of textbox, width, height
                    rgtBtmTf.Alignment = XParagraphAlignment.Center;
                    XPen pen = new XPen(XColors.Black, 1);
                    //rgtBtmGfx.DrawRectangle(pen, (page.Width - 240), (page.Height - 124), 237, 120);
                    rgtBtmGfx.DrawRectangle(pen, (page.Width - 200), (page.Height - 124), 197, 120);
                    rgtBtmTf.DrawString(rgtBtmWatermark, rgtFont, rgtBrush, rect, XStringFormats.TopLeft);
                    //rgtBtmTf.DrawString(rgtBtmWatermark, rgtFont, XBrushes.Blue, rect, XStringFormats.TopLeft);
                }

                // Save the document...
                //string opFilename = Path.GetFullPath("Data/Watermark_tempfile.pdf");
                //string FilePath = Server.MapPath("Watermark_tempfile.pdf");
                string opFilename = ipFilename; //"D:\\Sample_2.pdf";
                document.Save(opFilename);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        [HttpPost]
        public ActionResult PublishDocument()
        {
            string result = "";
            try
            {
                DocumentUpload bllOBJ = new DocumentUpload();
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.Action = "Published";
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                bool isDocumentUploaded = false;
                if (Request.Form["EditableDocumentUploaded"].ToString() == "yes")
                    isDocumentUploaded = true;
             
                if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                {
                    objDoc.ReadableDocumentName = Request.Form["EditableDocumentName"].ToString();
                }                    
                else
                {
                    objDoc.ReadableDocumentName = Request.Form["ReadableDocumentName"].ToString();
                }
                //save image to images folder
                HttpFileCollectionBase files = Request.Files;
                if(isDocumentUploaded)
                {
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
                            byte[] fileByteArray = new byte[file.ContentLength];
                            file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                            if (i == 0)
                            {
                                objDoc.EditableByteArray = fileByteArray;
                                objDoc.EditableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                                if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                                {
                                    objDoc.ReadableByteArray = fileByteArray;
                                    objDoc.ReadableDocumentName = objDoc.DocumentNo + Path.GetExtension(file.FileName);
                                }
                            }
                            else
                            {
                                objDoc.ReadableByteArray = fileByteArray;
                            }
                        }
                    }
                }
                else
                {
                    string EditableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName);
                    string ReadableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName);
                    if (Request.Form["DocumentCategoryCode"].ToString() == "FR")
                    {
                        ReadableURL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.ReadableDocumentName);
                    }
                    objDoc.EditableByteArray = bllOBJ.DownloadDocument(EditableURL);
                    objDoc.ReadableByteArray = bllOBJ.DownloadDocument(ReadableURL);
                }
                bllOBJ.DocumentPublish(objDoc, isDocumentUploaded);
                result = "sucess";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        //public ActionResult DownloadDraftDocument(string folderPath, double versionNo, string fileName)
        //{
        //    //string filename = "TEPL-COMMON-CE-XX-PO-0006.docx";
        //    //URL = @"file//D://Storage/QMS/DraftDocuments/CEO OFFICE/CEO OFFICE/POLICY/TEPL-COMMON-CE-XX-PO-0006.docx";
        //    try
        //    {
        //        //string URL = siteURL + "/" + docLib + "/" + folderPath + "/" + fileName;
        //        string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, folderPath, fileName);
        //        DocumentUpload bllOBJ = new DocumentUpload();
        //        byte[] fileContent = bllOBJ.DownloadDocument(URL);
        //        string contentType = MimeMapping.GetMimeMapping(fileName);

        //        var cd = new System.Net.Mime.ContentDisposition
        //        {
        //            FileName = fileName,
        //            Inline = true,
        //        };

        //        Response.AppendHeader("Content-Disposition", cd.ToString());

        //        return File(fileContent, contentType);
        //        //return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        LoggerBlock.WriteTraceLog(ex);
        //        return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
        //    }
        //}

        //public ActionResult DownloadReadableDocument(string folderPath, double versionNo, string fileName)
        //{
        //    try
        //    {
        //        string URL = CommonMethods.CombineUrl(QMSConstants.StoragePath, QMSConstants.DraftFolder, folderPath, fileName);
        //        DocumentUpload bllOBJ = new DocumentUpload();
        //        byte[] fileContent = bllOBJ.DownloadDocument(URL);

        //        return File(fileContent, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        LoggerBlock.WriteTraceLog(ex);
        //        return Json(new { success = true, message = "failed" }, JsonRequestBehavior.AllowGet);
        //    }
        //}


        [HttpPost]
        public ActionResult ConvertoPDF_Test()
        {
            string response = "";
            try
            {

            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                response = "error";
            }
            return Json(new { success = true, message = response }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ConvertToPDF()
        {
            string result = "";
            try
            {
                DraftDocument objDoc = CommonMethods.GetDocumentObject(Request.Form);
                objDoc.CompanyCode = QMSConstants.CompanyCode;
                objDoc.Action = "Published";
                objDoc.ActionedID = (Guid)System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID];
                objDoc.ActionByName = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserDisplayName].ToString();
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];
                    string flname; string temFileName = "";

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        flname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        flname = file.FileName;
                        temFileName = flname;
                        byte[] fileByteArray = new byte[file.ContentLength];
                        file.InputStream.Read(fileByteArray, 0, file.ContentLength);
                        objDoc.ReadableByteArray = fileByteArray;
                        string extension = Path.GetExtension(file.FileName);
                        objDoc.ReadableDocumentName = objDoc.DocumentNo + extension;
                        objDoc.ReadableFilePath = CommonMethods.CombineUrl(objDoc.ProjectName, objDoc.DepartmentName, objDoc.SectionName, objDoc.DocumentCategoryName);
                        objDoc.EditableFilePath = CommonMethods.CombineUrl(objDoc.ProjectName, objDoc.DepartmentName, objDoc.SectionName, objDoc.DocumentCategoryName);
                        FileConvert.ConvertToPDF(file.FileName, fileByteArray);
                    }
                }
                //Logic to convert document to PDF and send the path of link
                result = "/ApproveRequest/DownloadDocument?folderPath=NNB102%2FAdmin%2FADMIN%20COMMON%2FPolicies&versionNo=0.001&fileName=TEPL-NNB102-AD-AC-PO-0007.docx";
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                result = "failed";
                //throw ex;
            }
            return Json(new { success = true, message = result }, JsonRequestBehavior.AllowGet);
        }
    }
}