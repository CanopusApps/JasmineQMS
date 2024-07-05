using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using _Application = Microsoft.Office.Interop.Word._Application;

namespace TEPL.QMS.Common
{
    public static class FileConvert
    {
        public static byte[] ConvertToPDF(string strInputFile, byte[] inputByteArray)
        {
            //check the file type
            //convert file
            //encrypt file
            //download path
            //WordToPDF();
            //ExceltoPdf(@"C:\Siva\QMS Estimation.xlsx", @"C:\Siva\QMS Estimation.pdf");
            //powerPoint2Pdf(strInputFile, strOutputFile);
            string extension = Path.GetExtension(strInputFile);
            string outputFileName = strInputFile.Replace(extension, ".pdf");
            byte[] outputByteArray = null;
            switch (extension)
            {
                case ".docx":
                    outputByteArray = WordToPDF(strInputFile, outputFileName);
                    break;
                case ".doc":
                    outputByteArray = WordToPDF(strInputFile, outputFileName);
                    break;
                case ".xls":
                    ExceltoPdf(strInputFile, outputFileName);
                    break;
                case ".xlsx":
                    ExceltoPdf(strInputFile, outputFileName);
                    break;
                case ".ppt":
                    powerPoint2Pdf(strInputFile, outputFileName);
                    break;
                case ".pptx":
                    PPTXToPDF(strInputFile, outputFileName);
                    break;
                default:
                    break;
            }
            return outputByteArray;
        }
        //private static WordToPDF1()
        //{
        //    // Path to the input Word document
        //    string inputDocx = @"C:\input.docx";

        //    // Path to the output PDF file
        //    string outputPdf = @"C:\output.pdf";
            
        //    // Open the Word document for reading
        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(inputDocx, false))
        //    {
        //        // Create a new PDF document
        //        PdfDocument pdfDoc = new PdfDocument();

        //        // Add a new PDF page
        //        PdfPage pdfPage = pdfDoc.AddPage();

        //        // Create a PDF renderer based on the PDF page
        //        PdfRenderer renderer = new PdfRenderer(pdfPage);

        //        // Render the Word document to PDF
        //        renderer.Render(wordDoc);

        //        // Save the PDF document to disk
        //        pdfDoc.Save(outputPdf);
        //    }
        //}
        private static byte[] WordToPDF(string wordLocation,object outputFileName)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory
            FileInfo wordFile = new FileInfo(wordLocation);
            word.Visible = false;
            word.ScreenUpdating = false;
            string extension = Path.GetExtension(wordLocation);


            // Cast as Object for word Open method
            Object filename = (Object)wordLocation;

            // Use the dummy value as a placeholder for optional arguments
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            object fileFormat = WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;

            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            FileStream fs = new FileStream(outputFileName.ToString(),FileMode.Open);
            byte[] outByteArray = new byte[fs.Length];
            fs.Read(outByteArray, 0, outByteArray.Length);
            return outByteArray;
        }
        private static byte[] ExceltoPdf(string excelFileName, string outputFileName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open(excelFileName);
                wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFileName);

                wkb.Close();
                app.Quit();

                FileStream fs = new FileStream(outputFileName.ToString(), FileMode.Open);
                byte[] outByteArray = new byte[fs.Length];
                fs.Read(outByteArray, 0, outByteArray.Length);
                return outByteArray;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                throw ex;
            }
        }
        private static byte[] powerPoint2Pdf(string pptFileName, string outputFileName)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = null;
            Presentation pptPresentation = null;
            object unknownType = Type.Missing;
            try
            {
                //start power point   
                pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

                //open powerpoint document  
                pptPresentation = pptApplication.Presentations.Open(pptFileName, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                //export PDF from PPT  
                if (pptPresentation != null)
                {
                    pptPresentation.ExportAsFixedFormat((string)outputFileName,
                                                         PpFixedFormatType.ppFixedFormatTypePDF,
                                                         PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                         MsoTriState.msoFalse,
                                                         PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                                                         PpPrintOutputType.ppPrintOutputSlides,
                                                         MsoTriState.msoFalse, null,
                                                         PpPrintRangeType.ppPrintAll, string.Empty,
                                                         true, true, true, true, false, unknownType);
                 }
                else
                {
                    Console.WriteLine("Error occured for conversion of office PowerPoint to PDF");
                }
            }
            catch (Exception exPowerPoint2Pdf)
            {
                Console.WriteLine("Error occured for conversion of office PowerPoint to PDF, Exception: ", exPowerPoint2Pdf);
            }
            finally
            {
                // Close and release the Document object.  
                if (pptPresentation != null)
                {
                    pptPresentation.Close();
                    pptPresentation = null;
                }

                // Quit Word and release the ApplicationClass object.  
                pptApplication.Quit();
                pptApplication = null;
            }
            FileStream fs = new FileStream(outputFileName.ToString(), FileMode.Open);
            byte[] outByteArray = new byte[fs.Length];
            fs.Read(outByteArray, 0, outByteArray.Length);
            return outByteArray;
        }
        private static byte[] PPTXToPDF(string pptFileName, string outputFileName)
        {
            // Create COM Objects
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = null;
            Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation = null;
            try
            {
                object unknownType = Type.Missing;

                //start power point
                pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

                //open powerpoint document
                pptPresentation = pptApplication.Presentations.Open(pptFileName,
                    Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                // save PowerPoint as PDF
                pptPresentation.ExportAsFixedFormat(outputFileName,
                    Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                    Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    MsoTriState.msoFalse, Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    Microsoft.Office.Interop.PowerPoint.PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null,
                    Microsoft.Office.Interop.PowerPoint.PpPrintRangeType.ppPrintAll, string.Empty, true, true, true,
                    true, false, unknownType);
            }
            finally
            {
                // Close and release the Document object.
                if (pptPresentation != null)
                {
                    pptPresentation.Close();
                    pptPresentation = null;
                }

                // Quit PowerPoint and release the ApplicationClass object.
                if (pptApplication != null)
                {
                    pptApplication.Quit();
                    pptApplication = null;
                }
            }
            FileStream fs = new FileStream(outputFileName.ToString(), FileMode.Open);
            byte[] outByteArray = new byte[fs.Length];
            fs.Read(outByteArray, 0, outByteArray.Length);
            return outByteArray;
        }

    }
}
