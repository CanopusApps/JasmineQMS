using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TEPLQMS.Common
{
    public class LoggedInUser
    {
        public int ID { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string FullName { get; set; }
        public string Mobile { get; set; }
        public string Roles { get; set; }
    }
}