
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Web.Mvc;
using TEPL.QMS.Common;

namespace TEPLQMS.Controllers
{
    public class HomeController : Controller
    {
        TEPLQMS.Models.Database.DataBaseEntities db = new Models.Database.DataBaseEntities();

        [CustomAuthorize()]
        public ActionResult Index()
        {
            try
            {
                AddWatermarkonPDF("");
                LoggerBlock.WriteLog("In Home Controller and in try");
                //int x = 0;
                //int y = 5;
                //int z = y / x; 
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return View();
        }

        public void AddWatermarkonPDF(string pdfPath)
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
                string ipFilename = "D:\\Sample.pdf";

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
                string opFilename = "D:\\Sample_2.pdf";
                document.Save(opFilename);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public ActionResult Databoxes()
        {
            LoggerBlock.WriteLog("In Home Controller and in Databoxes view");
            return View();
        }

        public ActionResult DataTables()
        {
            LoggerBlock.WriteLog("In Home Controller and in DataTables view");
            return View();
        }

        public ActionResult AllControls()
        {
            return View();
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
