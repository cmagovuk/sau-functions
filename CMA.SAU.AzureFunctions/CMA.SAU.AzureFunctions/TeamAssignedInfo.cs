using System;
using System.Collections.Generic;
using System.Text;

namespace CMA.SAU.AzureFunctions
{
    internal class TeamAssignedInfo
    {
        public bool Changed { get; set; }
        public string RefNo { get; set; }
        public string Url { get; set; }
        public string Message { get; set; }
        public DateTime Created { get; set; }
    }
}
