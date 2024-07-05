using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class ExternalCategory
    {
        public Guid ID { get; set; }
        public string Title { get; set; }
        public bool Active { get; set; }
        public Guid CreatedBy { get; set; }
        public DateTime Created { get; set; }
        public Guid ModifiedBy { get; set; }
        public DateTime Modified { get; set; }
    }
}
