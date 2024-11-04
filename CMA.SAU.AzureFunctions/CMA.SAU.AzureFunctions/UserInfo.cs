using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CMA.SAU.AzureFunctions
{
    internal class UserInfo
    {
         public string Email { get; set; }
        public string Name { get; set; }
        public string First { get; set; }
        public string Surname { get; set; }
        public string Id { get; set; }
        public UserInfo(Microsoft.Graph.User user)
        {
            Email = user.Mail;
            Name = user.DisplayName;
            First = user.GivenName;
            Surname = user.Surname;
            Id = user.Id;
        }
    }
}
