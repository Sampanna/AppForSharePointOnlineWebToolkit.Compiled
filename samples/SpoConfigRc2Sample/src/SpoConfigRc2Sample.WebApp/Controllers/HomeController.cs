using AppForSharePointOnlineWebToolkit;

using Microsoft.AspNetCore.Mvc;

namespace SpoConfigRc2Sample.WebApp.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            var helper = new ClientContextHelper();
            var client = helper.CreateAppOnlyClientContext("https://[TENANT_NAME]-admin.sharepoint.com");
            client.Dispose();
            return this.View();
        }

        public IActionResult About()
        {
            this.ViewData["Message"] = "Your application description page.";

            return this.View();
        }

        public IActionResult Contact()
        {
            this.ViewData["Message"] = "Your contact page.";

            return this.View();
        }

        public IActionResult Error()
        {
            return this.View();
        }
    }
}
