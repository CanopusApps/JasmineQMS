using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Workflow.Models
{
    public class ApprovalMatrix
    {
        public Guid ID { get; set; }
        public Guid ProjectTypeID { get; set; }
        public string ProjectTypeName { get; set; }
        public Guid ProjectID { get; set; }
        public string ProjectName { get; set; }
        public Guid StageID { get; set; }
        public string StageName { get; set; }
        public string DocumentLevel { get; set; }
        public Guid DepartmentID { get; set; }
        public string DepartmentName { get; set; }
        public Guid SectionID { get; set; }
        public string SectionName { get; set; }
        public string ApprovalUsers { get; set; }
        public string ApprovalUserIDs { get; set; }
        public bool Active { get; set; }
        public Guid CreatedID { get; set; }
        public DateTime CreatedOn { get; set; }
        public Guid ModifiedID { get; set; }
        public DateTime ModifiedOn { get; set; }
    }
}
