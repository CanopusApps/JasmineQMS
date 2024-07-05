using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Workflow.Models
{
    public class WFUser
    {
        public Guid ID { get; set; }
        public string LoginID { get; set; }
        public string DisplayName { get; set; }
        public string EmailID { get; set; }
        public bool IsActive { get; set; }
    }
}
