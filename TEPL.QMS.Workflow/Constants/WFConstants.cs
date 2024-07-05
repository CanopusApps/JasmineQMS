using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
namespace TEPL.QMS.Workflow.Constants
{
    public static class WFConstants
    {
        public static string WFDBCon = Convert.ToString(ConfigurationManager.ConnectionStrings["WFDBCon"]);

        //Stored procedure
        public static string spWFInitiate = "spWFInitiate";
        public static string spCreateAction = "spCreateAction";
        public static string spExecuteAction = "spExecuteAction";
        public static string spGetWorkflowStage = "spGetWorkflowStage";
        public static string spGetPendingActionItems = "spGetPendingActionItems";
        public static string spGetWorkflowStages = "spGetWorkflowStages";
        public static string spGetExecutionDetails = "spGetExecutionDetails";
        public static string spGetWFStages = "spGetWFStages";
        public static string spGetWFApprovalMatrix = "spGetWFApprovalMatrix";
        public static string spGetWFApprovalMatrixForID = "spGetWFApprovalMatrixForID";
        public static string spCopyWFApprovalMatrix = "spCopyWFApprovalMatrix";
        public static string spGetWFApprover = "spGetWFApprover";
        public static string spGetWFApprovers = "spGetWFApprovers";
        public static string spInsertWFApprovalMatrix = "spInsertWFApprovalMatrix";
        public static string spUpdateWFApprovalMatrix = "spUpdateWFApprovalMatrix";
        public static string spUpdateWFApprovalMatrixApprovers = "spUpdateWFApprovalMatrixApprovers";
        public static string spDeleteWFApprovalMatrix = "spDeleteWFApprovalMatrix";
        public static string spDeleteWFApprovalMatrixItem = "spDeleteWFApprovalMatrixItem";
        public static string spDownloadWFApprovalMatrix = "spDownloadWFApprovalMatrix";
        public static string spGetApprovalMatrixForUser = "spGetApprovalMatrixForUser";
        public static string spGetUserApprovalItemsForReplace = "spGetUserApprovalItemsForReplace";
        public static string spUpdateUserApprovalItemsForReplace = "spUpdateUserApprovalItemsForReplace";
    }
}
