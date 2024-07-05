using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Models
{
    public class MultipleApproversDisplay
    {
        public List<User> DocumentReviewers { get; set; }
        public List<User> DocumentApprovers { get; set; }
    }
}
