using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class DocumentResponse
    {
        public string DocumentName { get; set; }
        public int DocVersion { get; set; }
        public string DocumentURL { get; set; }
    }
}
