using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Workflow.DAL;
using TEPL.QMS.Workflow.Log;
using TEPL.QMS.Workflow.Models;

namespace TEPL.QMS.Workflow.Business
{
    public class WFAdminBLL
    {
        DALWorkflowActions objDALWF = new DALWorkflowActions();
        WFAdminDAL objAdminDALWF = new WFAdminDAL();
        public DataTable GetWorkflowStages(Guid WorkflowID)
        {
            return objDALWF.GetWorkflowStages(WorkflowID);
        }
        public DataTable GetWFApprovalMatrix(Guid ProjectTypeID, Guid ProjectID, Guid StageID, string DocumentLevel, Guid DepartmentID, Guid SectionID)
        {
            return objAdminDALWF.GetWFApprovalMatrix(ProjectTypeID, ProjectID, StageID, DocumentLevel,DepartmentID,SectionID);
        }
        public DataTable GetWFApprovalMatrixForID(Guid ID)
        {
            return objAdminDALWF.GetWFApprovalMatrixForID(ID);
        }
        public string CopyWFApprovalMatrix(Guid sProjectID, Guid dProjectID, int syncType, Guid createdBy)
        {
            return objAdminDALWF.CopyWFApprovalMatrix(sProjectID, dProjectID, syncType, createdBy);
        }
        public DataTable GetUserApprovalItemsForReplace(string strCondition, Guid UserID)
        {
            DataTable dt = objAdminDALWF.GetUserApprovalItemsForReplace(strCondition, UserID);
            return dt;
        }
        public DataTable GetUserApprovalItemsForAMID(Guid AMID)
        {
            DataTable dt = objAdminDALWF.GetUserApprovalItemsForAMID(AMID);
            return dt;
        }
        public string UpdateUserApprovalItemsForReplace(string strCondition, Guid CurrentUserID, Guid NewUserID)
        {
            return objAdminDALWF.UpdateUserApprovalItemsForReplace(strCondition, CurrentUserID, NewUserID);
        }
        public string DeleteWFApprovalMatrix(Guid ID)
        {
            return objAdminDALWF.DeleteWFApprovalMatrix(ID);
        }
        public List<ApprovalMatrix> GetApprovalMatrixForUser(Guid UserID, Guid projectTypeID, Guid projectID)
        {
           //return objAdminDALWF.GetApprovalMatrixForUser(UserID);
            List<ApprovalMatrix> list = new List<ApprovalMatrix>();
            try
            {
                DataTable dt = objAdminDALWF.GetApprovalMatrixForUser(UserID, projectTypeID, projectID);
                
                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    ApprovalMatrix itm = new ApprovalMatrix();
                    itm.ID = new Guid(dt.Rows[z]["ID"].ToString());
                    if (dt.Rows[z]["ProjectTypeID"].ToString() != "") //00000000-0000-0000-0000-00000000000
                    {
                        itm.ProjectTypeID = new Guid(dt.Rows[z]["ProjectTypeID"].ToString());
                        itm.ProjectTypeName = dt.Rows[z]["ProjectTypeName"].ToString();
                    }
                    if (dt.Rows[z]["ProjectID"].ToString() != "") //00000000-0000-0000-0000-00000000000
                    {
                        itm.ProjectID = new Guid(dt.Rows[z]["ProjectID"].ToString());
                        itm.ProjectName = dt.Rows[z]["ProjectName"].ToString();
                    }
                    if (dt.Rows[z]["StageID"].ToString() != "")
                    {
                        itm.StageID = new Guid(dt.Rows[z]["StageID"].ToString());
                        itm.StageName = dt.Rows[z]["StageName"].ToString();
                    }
                    if (dt.Rows[z]["DocumentLevel"].ToString() != "")
                        itm.DocumentLevel = dt.Rows[z]["DocumentLevel"].ToString();
                    if (dt.Rows[z]["DepartmentID"].ToString() != "")
                    {
                        itm.DepartmentID = new Guid(dt.Rows[z]["DepartmentID"].ToString());
                        itm.DepartmentName = dt.Rows[z]["DepartmentName"].ToString();
                    }
                    if (dt.Rows[z]["SectionID"].ToString() != "")
                    {
                        itm.SectionID = new Guid(dt.Rows[z]["SectionID"].ToString());
                        itm.SectionName = dt.Rows[z]["SectionName"].ToString();
                    }
                    itm.ApprovalUserIDs = dt.Rows[z]["ApprovalUserIDs"].ToString();
                    itm.ApprovalUsers = dt.Rows[z]["ApprovalUsers"].ToString();
                    list.Add(itm);
                }
            }
            catch (Exception ex)
            {
                Common.LoggerBlock.WriteTraceLog(ex);
            }
            return list;
        }
        public string InsertWFApprovalMatrix(ApprovalMatrix objAM)
        {
            return objAdminDALWF.InsertWFApprovalMatrix(objAM);
        }
        public string UpdateWFApprovalMatrix(ApprovalMatrix objAM)
        {
            return objAdminDALWF.UpdateWFApprovalMatrix(objAM);
        }
        public string UpdateWFApprovalMatrixApprovers(ApprovalMatrix objAM)
        {
            return objAdminDALWF.UpdateWFApprovalMatrixApprovers(objAM);
        }
        public byte[] DownloadWFApprovalMatrix(Guid projectID, string tableName)
        {
            byte[] fileContent = null;
            try
            {
                string strExcelPath = QMSConstants.TempFolder + Guid.NewGuid() + ".xlsx";
                DataTable dt = objAdminDALWF.DownloadWFApprovalMatrix(projectID);
                string[] selectedColumns = new[] { "ProjectName", "StageName", "DocumentLevel", "DepartmentName", "SectionName", "ApprovalUsers" };
                DataTable dt1 = new DataView(dt).ToTable(false, selectedColumns);
                dt1.TableName = tableName;
                ExcelOperations.ExportDataSet(dt1, strExcelPath);
                fileContent = File.ReadAllBytes(strExcelPath);
            }
            catch(Exception ex)
            {
                Common.LoggerBlock.WriteTraceLog(ex);
            }
            return fileContent; 
        }
    }
}
