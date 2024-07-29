using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class Departments
    {
        public Guid ID { get; set; }
        public string Code { get; set; }
        public string Title { get; set; }
        public Guid HODID { get; set; }
        public string HODName { get; set; }
        public bool Active { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Created { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime Modified { get; set; }
        public Sections Sections { get; set; }
    }
}
