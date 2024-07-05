using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;

namespace TEPL.QMS.DAL.Database.Component
{
    public class TemplateDocumentDAL
    {
        public DataTable GetTemplateDocuments()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetTemplateDocument, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }
        public DataTable AddTemplate(TemplateDocument objTemp)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spAddTemplateDocument, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@TemplateCode", SqlDbType.NVarChar, 10).Value = objTemp.Code;
                        cmd.Parameters.Add("@TemplateName", SqlDbType.NVarChar, 100).Value = objTemp.Name;
                        cmd.Parameters.Add("@TemplateFileName", SqlDbType.NVarChar, 100).Value = objTemp.FileName;
                        cmd.Parameters.Add("@TemplateFilePath", SqlDbType.NVarChar, 100).Value = objTemp.FilePath;
                        cmd.Parameters.Add("@TemplateLevel", SqlDbType.NVarChar, 100).Value = objTemp.Level;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = objTemp.CreatedID;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }

        public DataTable UpdateTemplate(TemplateDocument objTemp)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spUpdateTemplate, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = objTemp.ID;
                        cmd.Parameters.Add("@txtName", SqlDbType.NVarChar, 100).Value = objTemp.Name;
                        cmd.Parameters.Add("@txtFileName", SqlDbType.NVarChar, 100).Value = objTemp.FileName;
                        //cmd.Parameters.Add("@txtFilePath", SqlDbType.NVarChar, 100).Value = objTemp.FilePath;
                        cmd.Parameters.Add("@txtLevel", SqlDbType.NVarChar, 100).Value = objTemp.Level;
                        cmd.Parameters.Add("@LastModifiedID", SqlDbType.UniqueIdentifier).Value = new Guid(objTemp.ModifiedBy);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }
    }
}
