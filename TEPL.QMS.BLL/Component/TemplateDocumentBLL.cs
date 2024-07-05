using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;
using TEPL.QMS.DAL.Database.Component;

namespace TEPL.QMS.BLL.Component
{
    public class TemplateDocumentBLL
    {
        TemplateDocumentDAL objTempDAL = new TemplateDocumentDAL();
        DAL.Database.Component.DocumentOperations docOperObj = new DAL.Database.Component.DocumentOperations();
        public List<TemplateDocument> GetTemplateDocuments()
        {
            List<TemplateDocument> list = new List<TemplateDocument>();
            try
            {
                DataTable dt = objTempDAL.GetTemplateDocuments();
                dt.DefaultView.Sort = "Name ASC";
                dt = dt.DefaultView.ToTable();

                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    TemplateDocument itm = new TemplateDocument();
                    itm.ID = new Guid(dt.Rows[z]["ID"].ToString());
                    itm.Code = dt.Rows[z]["Code"].ToString();
                    itm.Name = dt.Rows[z]["Name"].ToString();
                    itm.FileName = dt.Rows[z]["FileName"].ToString();
                    itm.FilePath = dt.Rows[z]["FilePath"].ToString();
                    itm.Level = dt.Rows[z]["Level"].ToString();
                    if (dt.Rows[z]["flgActive"].ToString().ToLower() == "true")
                        itm.Active = true;
                    else
                        itm.Active = false;
                    list.Add(itm);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return list;
        }
        public void AddTemplate(TemplateDocument objTemp)
        {
            try
            {
                docOperObj.UploadDocument(QMSConstants.StoragePath,QMSConstants.TemplateFolder, objTemp.FilePath, objTemp.FileName, 1, objTemp.byteArray);
                DataTable dt = objTempDAL.AddTemplate(objTemp);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
        }

        public void UpdateTemplate(TemplateDocument objTemp, bool isFileUploaded)
        {
            try
            {
                if (isFileUploaded)
                    docOperObj.UploadDocument(QMSConstants.StoragePath,QMSConstants.TemplateFolder, objTemp.FilePath, objTemp.FileName, 1, objTemp.byteArray);
                DataTable dt = objTempDAL.UpdateTemplate(objTemp);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
        }
    }
}
