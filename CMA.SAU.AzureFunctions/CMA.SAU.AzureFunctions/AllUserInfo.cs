using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CMA.SAU.AzureFunctions
{
    internal class AllUserInfo
    {
        public List<UserInfo> Admin { get; set; }
        public List<UserInfo> Lead { get; set; }
        public List<UserInfo> Team { get; set; }
    }
}
