using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Workflow.Models
{
    public class WFApprovers
    {
        public Guid ID { get; set; }
        public string StageName { get; set; }
        public List<WFCondition> Condition { get; set; }
        public string ApprovalUser { get; set; }
        public List<WFUser> WFUsers{ get; set; }
    }
    public class WFCondition
    {
        public string Level { get; set; }
        public string Section { get; set; }

    }
}
