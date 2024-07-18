using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class DocumentNumbers
    {
        public Guid ID { get; set; }
        public string DepartmentCode { get; set; }
        public string DepartmentName { get; set; }
        public string SectionCode { get; set; }
        public string SectionName { get; set; }
        public string ProjectCode { get; set; }
        public string ProjectName { get; set; }
        public string FunctionCode { get; set; }
        public string FunctionName { get; set; }
        public string DocumentCategoryCode { get; set; }
        public string DocumentCategoryName { get; set; }
        public string SerialNo { get; set; }
        public bool Active { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Created { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime Modified { get; set; }
    }
}
