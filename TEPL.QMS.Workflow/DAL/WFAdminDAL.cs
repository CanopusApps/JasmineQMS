using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Workflow.Constants;
using TEPL.QMS.Workflow.Log;
using TEPL.QMS.Workflow.Models;

namespace TEPL.QMS.Workflow.DAL
{
    public class WFAdminDAL
    {
        public string GetWFStages(Guid WorkflowID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWFStages, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = WorkflowID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;

        }
        public DataTable GetWFApprovalMatrix(Guid ProjectTypeID, Guid ProjectID, Guid StageID, string DocumentLevel, Guid DepartmentID, Guid SectionID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWFApprovalMatrix, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = ProjectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = ProjectID;
                        cmd.Parameters.Add("@StageID", SqlDbType.UniqueIdentifier).Value = StageID;
                        cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar,10).Value = DocumentLevel;
                        cmd.Parameters.Add("@DepartmentID", SqlDbType.UniqueIdentifier).Value = DepartmentID;
                        cmd.Parameters.Add("@SectionID", SqlDbType.UniqueIdentifier).Value = SectionID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        public DataTable GetWFApprovalMatrixForID(Guid ID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWFApprovalMatrixForID, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = ID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        public string CopyWFApprovalMatrix(Guid sProjectID, Guid dProjectID, int syncType, Guid createdBy)
        {
            DataTable dt = new DataTable();
            string result = "";
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spCopyWFApprovalMatrix, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@SProjectID", SqlDbType.UniqueIdentifier).Value = sProjectID;
                        cmd.Parameters.Add("@DProjectID", SqlDbType.UniqueIdentifier).Value = dProjectID;
                        cmd.Parameters.Add("@SyncType", SqlDbType.Int).Value = syncType;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = createdBy;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            result = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }
        public DataTable GetUserApprovalItemsForReplace(string strCondition, Guid UserID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetUserApprovalItemsForReplace, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Condition", SqlDbType.NVarChar,-1).Value = strCondition;
                        cmd.Parameters.Add("@UserID", SqlDbType.UniqueIdentifier).Value = UserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }
        public DataTable GetUserApprovalItemsForAMID(Guid AMID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetUserApprovalItemsForReplace, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@AMID", SqlDbType.UniqueIdentifier).Value = AMID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }
        public string UpdateUserApprovalItemsForReplace(string strCondition, Guid CurrentUserID,Guid NewUserID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spUpdateUserApprovalItemsForReplace, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Condition", SqlDbType.NVarChar, -1).Value = strCondition;
                        cmd.Parameters.Add("@CurrentUserID", SqlDbType.UniqueIdentifier).Value = CurrentUserID;
                        cmd.Parameters.Add("@NewUserID", SqlDbType.UniqueIdentifier).Value = NewUserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }

        public DataTable GetApprovalMatrixForUser(Guid UserID, Guid projectTypeID, Guid projectID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetApprovalMatrixForUser, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@UserID", SqlDbType.UniqueIdentifier).Value = UserID;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = projectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = projectID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                       }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        public string InsertWFApprovalMatrix(ApprovalMatrix objAM)
        {
            DataTable dt = new DataTable();
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spInsertWFApprovalMatrix, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = objAM.ProjectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = objAM.ProjectID;
                        cmd.Parameters.Add("@StageID", SqlDbType.UniqueIdentifier).Value = objAM.StageID;
                        cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar, 10).Value = objAM.DocumentLevel;
                        cmd.Parameters.Add("@SectionID", SqlDbType.UniqueIdentifier).Value = objAM.SectionID;
                        cmd.Parameters.Add("@Users", SqlDbType.NVarChar, -1).Value = objAM.ApprovalUserIDs;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = objAM.CreatedID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }
        public string UpdateWFApprovalMatrix(ApprovalMatrix objAM)
        {
            DataTable dt = new DataTable();
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spUpdateWFApprovalMatrix, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = objAM.ID;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = objAM.ProjectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = objAM.ProjectID;
                        cmd.Parameters.Add("@StageID", SqlDbType.UniqueIdentifier).Value = objAM.StageID;
                        cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar, 10).Value = objAM.DocumentLevel;
                        cmd.Parameters.Add("@SectionID", SqlDbType.UniqueIdentifier).Value = objAM.SectionID;
                        cmd.Parameters.Add("@Users", SqlDbType.NVarChar, -1).Value = objAM.ApprovalUserIDs;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }

        public string UpdateWFApprovalMatrixApprovers(ApprovalMatrix objAM)
        {
            DataTable dt = new DataTable();
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spUpdateWFApprovalMatrixApprovers, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = objAM.ID;
                        cmd.Parameters.Add("@Users", SqlDbType.NVarChar, -1).Value = objAM.ApprovalUserIDs;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }

        public string DeleteWFApprovalMatrix(Guid ID)
        {
            DataTable dt = new DataTable();
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spDeleteWFApprovalMatrixItem, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = ID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }

        public DataTable DownloadWFApprovalMatrix(Guid projectID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spDownloadWFApprovalMatrix, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = projectID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

    }
}
