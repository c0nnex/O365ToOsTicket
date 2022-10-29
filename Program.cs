// Copyright (c) Ulrich Strauss. All rights reserved.
// Licensed under the MIT License.

using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace O365ToOsTicket
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                await RunAsync();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        private static AuthenticationConfig config;
        private static async Task RunAsync()
        {
            config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            // You can run this sample using ClientSecret or Certificate. The code will differ only when instantiating the IConfidentialClientApplication
            bool isUsingClientSecret = IsAppUsingClientSecret(config);

            // Even if this is a console application here, a daemon application is a confidential client application
            IConfidentialClientApplication app;

            if (isUsingClientSecret)
            {
                // Even if this is a console application here, a daemon application is a confidential client application
                app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                    .WithClientSecret(config.ClientSecret)
                    .WithAuthority(new Uri(config.Authority))
                    .Build();
            }

            else
            {
                ICertificateLoader certificateLoader = new DefaultCertificateLoader();
                certificateLoader.LoadIfNeeded(config.Certificate);

                app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                    .WithCertificate(config.Certificate.Certificate)
                    .WithAuthority(new Uri(config.Authority))
                    .Build();
            }

            app.AddInMemoryTokenCache();

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator. 
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; 
            await CallMSGraphUsingGraphSDK(app, scopes);

        }


        private static async Task CallMSGraphUsingGraphSDK(IConfidentialClientApplication app, string[] scopes)
        {
            // Prepare an authenticated MS Graph SDK client
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient(app, scopes);
            try
            {
                var folders = await graphServiceClient.Users[config.EmailAddress].MailFolders.Request().GetAsync();
                int i = 0;
                foreach (var folder in folders)
                {
                    // Console.WriteLine($"{folder.DisplayName}:{folder.TotalItemCount}");
                    if (folder.DisplayName == config.InboxName)
                    {
                        var msgs = await graphServiceClient.Users[config.EmailAddress].MailFolders[folder.Id].Messages.Request().Expand("attachments").GetAsync();
                        foreach (var msg in msgs)
                        {
                            Console.WriteLine($"{msg.From.EmailAddress.Address} {msg.ReceivedDateTime} {msg.Attachments.Count()} {msg.Subject}");
                            foreach (var item in msg.Attachments)
                            {
                                Console.WriteLine($"   {item.Name} {item.Size} {item.ContentType}");
                            }
                            var mtmsg = await graphServiceClient.Users[config.EmailAddress].Messages[msg.Id].Content.Request().GetAsync();
                            var mContent = new StreamReader(mtmsg).ReadToEnd();

                            /*    
                            - Call https://osticket/api/tickets.email and pipe email into it.
                            - Pipe email into php api/pipe.php
                            */

                            var osTicketResult = await PostToOSTicket(mContent);
                            if (osTicketResult != HttpStatusCode.Created)
                            {
                                Console.WriteLine("Problem creating ticket. Bailing out");
                                return;
                            }
#if DUMPINPUT
                            var path = (i++) + ".txt";
                            using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
                            {
                                mtmsg.CopyTo(outputFileStream);
                            }

#endif
                            if (config.DeleteProcessed)
                                await graphServiceClient.Users[config.EmailAddress].Messages[msg.Id].Request().DeleteAsync();

                        }
                    }
                }
            }
            catch (ServiceException e)
            {
                Console.WriteLine("Could not process emails: " + e);
            }

        }


        private static async Task<HttpStatusCode> PostToOSTicket(string data)
        {
            try
            {
                RestClient client = new RestClient(config.OsTicketUrl);

                RestRequest request = new RestRequest($"/api/tickets.email", Method.Post);
                //request.JsonSerializer = new RestSharp.Serializers.NetwonsoftJsonSerializer();
                request.AddHeader("X-API-KEY", config.OsTicketApiKey);
                request.RequestFormat = DataFormat.None;
                request.AddBody(data);



                Console.WriteLine("Sending...");
                var response = await client.ExecuteAsync(request, CancellationToken.None);
                Console.WriteLine($"StatusCode: {response.StatusCode}, Content-Type: {response.ContentType}, Content-Length: {response.ContentLength}):\r\n{response.Content}");
                return response.StatusCode;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
                return HttpStatusCode.InternalServerError;
            }
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
        {

            GraphServiceClient graphServiceClient =
                    new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                        AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));

            return graphServiceClient;
        }




        /// <summary>
        /// Checks if the sample is configured for using ClientSecret or Certificate. This method is just for the sake of this sample.
        /// You won't need this verification in your production application since you will be authenticating in AAD using one mechanism only.
        /// </summary>
        /// <param name="config">Configuration from appsettings.json</param>
        /// <returns></returns>
        private static bool IsAppUsingClientSecret(AuthenticationConfig config)
        {
            string clientSecretPlaceholderValue = "[Enter here a client secret for your application]";

            if (!String.IsNullOrWhiteSpace(config.ClientSecret) && config.ClientSecret != clientSecretPlaceholderValue)
            {
                return true;
            }

            else if (config.Certificate != null)
            {
                return false;
            }

            else
                throw new Exception("You must choose between using client secret or certificate. Please update appsettings.json file.");
        }
    }
}
