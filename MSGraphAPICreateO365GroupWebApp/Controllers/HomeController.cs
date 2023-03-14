using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MSGraphAPICreateO365GroupWebApp.Models;
using System.Diagnostics;
using Microsoft.Graph.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.VisualBasic.FileIO;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.Graph.Models.ODataErrors;

namespace MSGraphAPICreateO365GroupWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;
        private GraphServiceClient _graphClient;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;

            string tenantId = _configuration.GetValue<string>("AzureAd:TenantId");
            string clientId = _configuration.GetValue<string>("AzureAd:ClientId");
            string clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");
            string[] scopes = { "https://graph.microsoft.com/.default" };

            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);

            GraphServiceClient graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

            _graphClient = graphServiceClient;
        }

        #region List groups
        public async Task<IActionResult> Index()
        {
            
            try
            {
                GroupCollectionResponse? groups = await _graphClient.Groups.GetAsync();
                if (groups != null)
                {
                    ViewData["GraphApiResult"] = groups.Value;

                    ViewData["GroupCount"] = groups.OdataCount.ToString();
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
        
        /// <summary>
        /// A function that creates a group in Office 365.
        /// </summary>
        /// <returns>
        /// The view is being returned.
        /// </returns>
        public ActionResult CreateO365Group()
        {
            if (!ModelState.IsValid) return new StatusCodeResult(404);
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public async Task<ActionResult> CreateO365Group(IFormCollection collection)
        {
            try
            {
                string displayName = collection["groupName"];
                string description = collection["description"];
                string[] resourceBehaviorOptions = collection["resourceBehaviorOptions"];

                var group = new Group
                {
                    Description = description,
                    DisplayName = displayName,
                    GroupTypes = new List<string>()
                    {
                        "Unified"
                    },
                    MailEnabled = true,
                    MailNickname = displayName,
                    SecurityEnabled = false,
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "owners@odata.bind" , new List<string>
                            {
                                "https://graph.microsoft.com/v1.0/users/b1b1296c-6d93-47ae-bc56-1ccbd000f0e2",
                            }
                        },
                        {
                            "resourceBehaviorOptions" , new List<string>(resourceBehaviorOptions)
                        },
                    },
                };

                var resultGroup = await _graphClient.Groups
                    .PostAsync(group);
                
                return RedirectToAction("Index");
            }
            catch (ODataError ex)
            {
                Console.WriteLine(ex.Error?.Message);
                return RedirectToAction("Error", "Home");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        public async Task<ActionResult> EditO365Group(string id)
        {
            try
            {
                if (id == null)
                {
                    return new StatusCodeResult(404);
                }
                var groupInfoTask = _graphClient.Groups[id].GetAsync();
                var getMembersTask = _graphClient.Groups[id].Members.GetAsync();
                await Task.WhenAll(groupInfoTask, getMembersTask);

                var groupInfo = await groupInfoTask;
                var members = await getMembersTask;

                if (members != null)
                {
                    ViewBag.memberscount = members.OdataCount;
                    ViewBag.members = members.Value;
                }
                
                return View(groupInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        [HttpPost]
        public async Task<ActionResult> EditO365Group(IFormCollection collection)
        {
            try
            {
                var group = new Group
                {
                    Description = collection["description"],
                    DisplayName = collection["displayname"],
                    GroupTypes = new List<string>()
                    {
                        collection["type"]
                    },
                    MailEnabled = true,
                    Mail = collection["mail"]
                };

                await _graphClient.Groups[collection["id"]]
                    .PatchAsync(group);
                
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }

        }

        public async Task<ActionResult> DeleteO365Group(string id)
        {
            try
            {
                if (id == null)
                {
                    return new StatusCodeResult(404);
                }
                var rowToEdit = await _graphClient.Groups[id].GetAsync();
                return View(rowToEdit);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        [HttpPost]
        public async Task<ActionResult> DeleteO365Group(IFormCollection collection)
        {
            try
            {
                string groupId = collection["id"];
                string displayName = collection["displayname"];
                await _graphClient.Groups[groupId]
                    .DeleteAsync();
                
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        public async Task<ActionResult> AddUserToGroup()
        {
            try
            {
                var listGroup = await _graphClient.Groups.GetAsync();
                ViewBag.listGroup = listGroup;
                return View();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }
        #endregion

        #region user and group
        
        public async Task<ActionResult> RemoveUserFromGroup(string userId, string groupId, string groupDisplayName)
        {
            try
            {
                var user = await _graphClient.Users[userId].GetAsync();

                O365UserGroup o365UserGroup = new O365UserGroup()
                {
                    GroupId = groupId,
                    UserId = userId,
                    UserDisplayName = user.DisplayName,
                    GroupDisplayName = groupDisplayName
                };
                return View(o365UserGroup);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        [HttpPost]
        public async Task<ActionResult> RemoveUserFromGroup(IFormCollection collection)
        {
            try
            {
                string groupId = collection["groupid"];
                string userId = collection["userid"];

                await _graphClient.Groups[groupId].Members[userId].Ref.DeleteAsync();
                
                return RedirectToAction("EditO365Group", new { id = groupId });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        
        public ActionResult CheckUserInGroup()
        {
            try
            {
                return View();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }

        public async Task<ActionResult> CheckUserInGroup(string acc)
        {
            try
            {
                string userPrincipalName = acc + '@' + O365Connection.Instant.O365Domain;

                var user = await _graphClient.Users[userPrincipalName].GetAsync();


                return View();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex?.Message);
                return RedirectToAction("Error", "Home");
            }
        }
        #endregion

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}