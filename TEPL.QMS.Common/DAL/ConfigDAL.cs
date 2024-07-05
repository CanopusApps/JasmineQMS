using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common.Constants;

namespace TEPL.QMS.Common.DAL
{
    public static class ConfigDAL
    {
        public static string GetConfigValue(string strParamKey)
        {
            string strValue = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetConfigValue, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ParamKey", SqlDbType.NVarChar, 100).Value = strParamKey;

                        var ParamValue = cmd.Parameters.Add("@ParamValue", SqlDbType.NVarChar, -1);
                        ParamValue.Direction = ParameterDirection.Output;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        strValue = (string)ParamValue.Value;
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strValue;
        }

        public static DataTable GetTaskSchedules()
        {
            DataTable dt=new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetTaskSchedules, con))
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
    }
}
