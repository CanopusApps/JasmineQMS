using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Workflow.Constants;
using TEPL.QMS.Workflow.Log;

namespace TEPL.QMS.Workflow.DAL
{
    public class DALWorkflowActions
    {
        public DataTable GetWorkflowStages(Guid WorkflowID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWorkflowStages, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = WorkflowID;
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
        public DataTable GetWorkflowStage(Guid WorkflowID, Guid CurrentStageID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWorkflowStage, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = WorkflowID;
                        cmd.Parameters.Add("@CurrentStageID", SqlDbType.UniqueIdentifier).Value = CurrentStageID;
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
        public DataTable GetWorkflowApprover(Guid ProjectTypeID,Guid ProjectID, Guid StageID, string DocumentLevel, Guid SectionID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWFApprover, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = ProjectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = ProjectID;
                        cmd.Parameters.Add("@StageID", SqlDbType.UniqueIdentifier).Value = StageID;
                        cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar, 10).Value = DocumentLevel;
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
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }
        public string GetWorkflowApprovers(Guid WorkflowID,Guid ProjectTypeID, Guid ProjectID, string DocumentLevel, Guid SectionID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetWFApprovers, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = WorkflowID;
                        cmd.Parameters.Add("@ProjectTypeID", SqlDbType.UniqueIdentifier).Value = ProjectTypeID;
                        cmd.Parameters.Add("@ProjectID", SqlDbType.UniqueIdentifier).Value = ProjectID;
                        cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar, 10).Value = DocumentLevel;
                        cmd.Parameters.Add("@SectionID", SqlDbType.UniqueIdentifier).Value = SectionID;
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
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public Guid WorkflowInitate(Guid WorkflowID, Guid DocumentID, Guid WorkflowStageID, Guid ActionedID, Guid CreatedID)
        {
            Guid InstanceId = new Guid();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spWFInitiate, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = WorkflowID;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = DocumentID;
                        cmd.Parameters.Add("@WorkflowStageID", SqlDbType.UniqueIdentifier).Value = WorkflowStageID;
                        cmd.Parameters.Add("@ActionedID", SqlDbType.UniqueIdentifier).Value = ActionedID;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = CreatedID;
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
            return InstanceId;
        }
        public string CreateAction(Guid ExecutionID, Guid WorkflowStageID, string ActionBy,string MultipleApprovers, Guid CreatedBy)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spCreateAction, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ExecutionID", SqlDbType.UniqueIdentifier).Value = ExecutionID;
                        cmd.Parameters.Add("@WorkflowStageID", SqlDbType.UniqueIdentifier).Value = WorkflowStageID;
                        cmd.Parameters.Add("@ActionedID", SqlDbType.NVarChar, -1).Value = ActionBy;
                        cmd.Parameters.Add("@MultipleApprovers", SqlDbType.NVarChar, -1).Value = MultipleApprovers;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = CreatedBy;

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
            return strReturn;
        }
        public string ExecuteAction(Guid ExecutionID, Guid WorkflowStageID, Guid ActionedID, string WorkflowAction, string ActionComments, Guid CreatedID, bool isDocumentUploaded)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spExecuteAction, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ExecutionID", SqlDbType.UniqueIdentifier).Value = ExecutionID;
                        cmd.Parameters.Add("@WorkflowStageID", SqlDbType.UniqueIdentifier).Value = WorkflowStageID;
                        cmd.Parameters.Add("@ActionedID", SqlDbType.UniqueIdentifier).Value = ActionedID;
                        cmd.Parameters.Add("@WorkflowAction", SqlDbType.NVarChar, 50).Value = WorkflowAction;
                        cmd.Parameters.Add("@ActionComments", SqlDbType.NVarChar, -1).Value = ActionComments;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = CreatedID;
                        cmd.Parameters.Add("@DocsUploaded", SqlDbType.Bit).Value = isDocumentUploaded;
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
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public DataTable GetExecutionDetails(Guid ExecutionID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetExecutionDetails, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ExecutionID", SqlDbType.UniqueIdentifier).Value = ExecutionID;
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
        public DataTable GetPendingActions(Guid ActionedID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(WFConstants.WFDBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(WFConstants.spGetPendingActionItems, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.SelectCommand.Parameters.Add("@ActionedID", SqlDbType.UniqueIdentifier).Value = ActionedID;
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
