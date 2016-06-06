using System.Collections.Generic;

using Microsoft.AspNet.Http.Authentication;
using Microsoft.AspNet.Identity;

namespace SpoConfigRc1Sample.WebApp.ViewModels.Manage
{
    public class ManageLoginsViewModel
    {
        public IList<UserLoginInfo> CurrentLogins { get; set; }

        public IList<AuthenticationDescription> OtherLogins { get; set; }
    }
}
