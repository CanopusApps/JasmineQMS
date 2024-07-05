using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class Project
    {
        public Guid ID { get; set; }
        public string Code { get; set; }
        public string Title { get; set; }
        public Guid ProjectTypeID { get; set; }
        public string ProjectTypeName { get; set; }
        public string ProjectTypeCode { get; set; }
        public Guid WorkflowID { get; set; }
        public string WorkflowName { get; set; }
        public bool IsAdmin { get; set; }
        public bool ProjectActive { get; set; }
        public Guid CreatedID { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Created { get; set; }
        public Guid ModifiedID { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime Modified { get; set; }
    }
}
