using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class User
    {
        public Guid ID { get; set; }
        public string LoginID { get; set; }
        public string DisplayName { get; set; }
        public string EmailID { get; set; }
        public string DepartmentName { get; set; }
        public bool IsProjectAdmin { get; set; }
        public bool IsQMSAdmin { get; set; }
        public Guid ProjectTypeID { get; set; }
        public Guid? ProjectID { get; set; }
        public bool IsActive { get; set; }
    }
}
