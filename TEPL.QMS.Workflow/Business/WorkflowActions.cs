using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common;
using TEPL.QMS.Workflow.DAL;
using TEPL.QMS.Workflow.Log;
using TEPL.QMS.Workflow.Models;

namespace TEPL.QMS.Workflow.Business
{
    public class WorkflowActions
    {
        DALWorkflowActions objDALWF = new DALWorkflowActions();
        WFAdminDAL objAdminDALWF = new WFAdminDAL();
        public DataTable GetWorkflowStages(Guid WorkflowID)
        {
            return objDALWF.GetWorkflowStages(WorkflowID);
        }
        public List<WFStage> GetWFStages(Guid WorkflowID)
        {
            List<WFStage> list = new List<WFStage>();
            string strReturn = objAdminDALWF.GetWFStages(WorkflowID);
            list = BindModels.ConvertJSON<WFStage>(strReturn);
            return list;
        }

        public List<DocumentApprover> GetWorkflowApprover(Guid ProjectTypeID, Guid ProjectID, Guid NextStageID, string DocumentLevel, Guid SectionID)
        {
            List<DocumentApprover> objApprovers = null;
            DataTable dt = objDALWF.GetWorkflowApprover(ProjectTypeID, ProjectID, NextStageID, DocumentLevel, SectionID);
            objApprovers = BindModels.ConvertDataTable<DocumentApprover>(dt);
            return objApprovers;
        }
        public List<WFApprovers> GetWorkflowApprovers(Guid WorkflowID, Guid ProjectTypeID, Guid ProjectID, string DocumentLevel, Guid SectionID)
        {
            List<WFApprovers> objWFApprovers = null;
            string strReturn = objDALWF.GetWorkflowApprovers(WorkflowID, ProjectTypeID, ProjectID, DocumentLevel, SectionID);
            objWFApprovers = BindModels.ConvertJSON<WFApprovers>(strReturn);
            return objWFApprovers;
        }
        public Stage GetWorkflowStage(Guid WorkflowID, Guid CurrentStageID)
        {
            DataTable dt = objDALWF.GetWorkflowStage(WorkflowID, CurrentStageID);
            List<Stage> objSt = new List<Stage>();
            objSt = BindModels.ConvertDataTable<Stage>(dt);

            return objSt[0];
        }

        public Guid WorkflowInitate(Guid WorkflowID, Guid DocumentID, Guid WFStageID, Guid ActionedID, Guid CreatedID)
        {
            return objDALWF.WorkflowInitate(WorkflowID, DocumentID, WFStageID, ActionedID, CreatedID);
        }
        public string ApproveAction(Guid WorkflowID, Guid ExecutionID, Guid WorkflowStageID, Guid NextWorkflowStageID, Guid ActionedID, string NextActionedID, string WorkflowAction, string ActionComments, string MulitpleApprovers, Guid CreatedID)
        {
            ExecuteAction(ExecutionID, WorkflowStageID, ActionedID, WorkflowAction, ActionComments, CreatedID, false);
            Stage objSt = GetWorkflowStage(WorkflowID, WorkflowStageID);
            CreateAction(ExecutionID, NextWorkflowStageID, NextActionedID, MulitpleApprovers, CreatedID);
            return "";
        }
        public string CreateAction(Guid ExecutionID, Guid WorkflowStageID, string ActionedID, string MultipleApprovers, Guid CreatedID)
        {
            return objDALWF.CreateAction(ExecutionID, WorkflowStageID, ActionedID, MultipleApprovers, CreatedID);
        }
        public string ExecuteAction(Guid ExecutionID, Guid WorkflowStageID, Guid ActionedID, string WorkflowAction, string ActionComments, Guid CreatedID, bool isDocumentUploaded)
        {
            return objDALWF.ExecuteAction(ExecutionID, WorkflowStageID, ActionedID, WorkflowAction, ActionComments, CreatedID, isDocumentUploaded);
        }
        public DataTable GetExecutionDetails(Guid ExecutionID)
        {
            return objDALWF.GetExecutionDetails(ExecutionID);
        }
        public DataTable GetPendingActions(Guid ActionedID)
        {
            return objDALWF.GetPendingActions(ActionedID);
        }
    }
}
