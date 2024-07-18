using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class DraftDocument
    {
        public int SeqNo { get; set; }
        public int RNo { get; set; }
        public Guid DocumentID { get; set; }
        public Guid DocumentPublishID { get; set; }
        public string CompanyCode { get; set; }
        public string DocumentNo { get; set; }
        public string DocumentDescription { get; set; }
        public string RevisionReason { get; set; }
        public Guid DepartmentID { get; set; }
        public string DepartmentCode { get; set; }
        public string DepartmentName { get; set; }
        public string DocumentLevel { get; set; }
        public Guid SectionID { get; set; }
        public string SectionCode { get; set; }
        public string SectionName { get; set; }
        public Guid ProjectID { get; set; }
        public string ProjectCode { get; set; }
        public string ProjectName { get; set; }
        public Guid ProjectTypeID { get; set; }
        public string ProjectTypeCode { get; set; }
        public string ProjectTypeName { get; set; }
        public Guid DocumentCategoryID { get; set; }
        public string DocumentCategoryCode { get; set; }
        public string DocumentCategoryName { get; set; }
        public Guid FunctionID { get; set; }
        public string FunctionCode { get; set; }
        public string FunctionName { get; set; }
        public string EditableDocumentName { get; set; }
        public string EditableFilePath { get; set; }
        public byte[] EditableByteArray { get; set; }
        public string ReadableDocumentName { get; set; }
        public string ReadableFilePath { get; set; }
        public byte[] ReadableByteArray { get; set; }
        public decimal DraftVersion { get; set; }
        public decimal OriginalVersion { get; set; }
        public decimal EditVersion { get; set; }
        public int Version { get; set; }
        public string Comments { get; set; }
        public bool IsMultipleApprovers { get; set; }
        public MultipleApprovers MultipleApprovers { get; set; }
        public MultipleApproversDisplay MultipleApproversDisplay { get; set; }
        public Guid WorkflowID { get; set; }
        public Guid WFExecutionID { get; set; }
        public Guid ActionedID { get; set; }
        public string ActionByName { get; set; }
        public string Action { get; set; }
        public string ActionComments { get; set; }
        public Guid UploadedUserID { get; set; }
        public string UploadedUserName { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string PublishedOn { get; set; }
        public Guid CurrentStageID { get; set; }
        public string CurrentStage { get; set; }
        public string WFStatus { get; set; }
        public int PendingDays { get; set; }

        public List<WorkflowStages> WorkflowStages;
    }

    public class WorkflowStages
    {
        public int RNo { get; set; }
        public Guid ID { get; set; }
        public string StageName { get; set; }
        public Guid WFActionedID { get; set; }
        public bool WFActionMandatory { get; set; }
        public string ActionUserName { get; set; }
        public string WFAction { get; set; }
        public DateTime? WFActionDate { get; set; }
        public string ActionComments { get; set; }
        public string Status { get; set; }
    }
}
