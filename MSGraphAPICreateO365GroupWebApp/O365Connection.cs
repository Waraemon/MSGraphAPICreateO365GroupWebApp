using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSGraphAPICreateO365GroupWebApp
{
    public class O365Connection
    {
        // requires using Microsoft.Extensions.Configuration;
        static readonly IConfiguration Configuration;
        static string tenantId = Configuration["ida:TenantId"];
        static string clientId = Configuration["ida:ClientId"];
        static string clientSecret = Configuration["ida:ClientSecret"];
        static string o365Domain = Configuration["ida:Domain"];
        private O365Connection()
        {

        }
        private static O365Connection instant { get; set; }
        public static O365Connection Instant
        {
            get
            {
                if (instant == null)
                {
                    instant = new O365Connection();
                    string[] scopes = { "https://graph.microsoft.com/.default" };

                    ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);

                    instant.graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);
                }
                return instant;
            }
        }

        private GraphServiceClient graphServiceClient { get; set; }
        public GraphServiceClient GraphServiceClient
        {
            get
            {
                return graphServiceClient;
            }
        }

        public string O365Domain
        {
            get
            {
                return o365Domain;
            }
        }
    }
}
