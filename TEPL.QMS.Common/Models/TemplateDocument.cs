using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class TemplateDocument
    {
        public Guid ID { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Level { get; set; }
        public byte[] byteArray { get; set; }
        public bool Active { get; set; }
        public Guid CreatedID { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Created { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime Modified { get; set; }
    }
}
