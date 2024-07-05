using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Model
{
    public class Requests
    {
        public string DocumentNo { get; set; }
        public string Category { get; set; }
        public string Department { get; set; }
        public string Section { get; set; }
        public string UploadedBy { get; set; }
        public DateTime UploadedOn { get; set; }
        public string Stage { get; set; }
        public string PendingWith { get; set; }
    }
}
