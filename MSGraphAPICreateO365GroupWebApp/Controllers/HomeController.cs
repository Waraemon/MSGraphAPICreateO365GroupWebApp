using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MSGraphAPICreateO365GroupWebApp.Models;
using System.Diagnostics;
using Microsoft.Graph.Models;

namespace MSGraphAPICreateO365GroupWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public async Task<IActionResult> Index()
        {
            string tenantId = _configuration.GetValue<string>("AzureAd:TenantId");
            string clientId = _configuration.GetValue<string>("AzureAd:ClientId");
            string clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");
            string[] scopes = { "https://graph.microsoft.com/.default" };

            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);

            GraphServiceClient graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);
            try
            {
                GroupCollectionResponse? groups = await graphServiceClient.Groups.GetAsync();
                if (groups != null)
                {
                    ViewData["GraphApiResult"] = groups.Value;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}