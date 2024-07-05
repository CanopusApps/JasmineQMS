using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class LoginUser
    {
        public Guid ID { get; set; }
        public string LoginID { get; set; }
        public string DisplayName { get; set; }
        public string EmailID { get; set; }
        public string Roles { get; set; }
        public string DepartmentName { get; set; }
        public string CreatedBy { get; set; }
        public List<Project> Projects { get; set; }
    }
}
