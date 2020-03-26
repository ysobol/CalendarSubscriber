using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using msgraphapp.Models;
using Newtonsoft.Json;
using System.Net;
using System.Threading;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;

namespace msgraphapp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NotificationsController : ControllerBase
    {
        private static Dictionary<string, Subscription> Subscriptions = new Dictionary<string, Subscription>();
        private static Timer subscriptionTimer = null;
        private readonly MyConfig config;
        private static object DeltaLink = null;
        private static IUserDeltaCollectionPage lastPage = null;

        public NotificationsController(MyConfig config)
        {
            this.config = config;
        }

        private async Task CheckForUpdates()
        {
            var graphClient = GetGraphClient();
            var users = await GetUsers(graphClient, DeltaLink); // get a page of users
            OutputUsers(users);
            // go through all of the pages so that we can get the delta link on the last page.
            while (users.NextPageRequest != null)
            {
                users = users.NextPageRequest.GetAsync().Result;
                OutputUsers(users);
            }

            object deltaLink;

            if (users.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLink))
            {
                DeltaLink = deltaLink;
            }
        }

        private void OutputUsers(IUserDeltaCollectionPage users)
        {
            users.ToList().ForEach(user => Console.WriteLine($"User: {user.Id}, {user.GivenName} {user.Surname}"));
        }

        private async Task<IUserDeltaCollectionPage> GetUsers(GraphServiceClient graphClient, object deltaLink)
        {
            IUserDeltaCollectionPage page;

            if (lastPage == null)
            {
                page = await graphClient.Users
                    .Delta()
                    .Request()
                    .GetAsync();
            }
            else
            {
                lastPage.InitializeNextPageRequest(graphClient, deltaLink.ToString());
                page = await lastPage.NextPageRequest.GetAsync();
            }

            lastPage = page;
            return page;
        }

        private void CheckSubscriptions(Object stateInfo)
        {
            AutoResetEvent autoEvent = (AutoResetEvent)stateInfo;
            Console.WriteLine($"Checking subscriptions {DateTime.Now.ToString("h:mm:ss.fff")}");
            Console.WriteLine($"Current subscription count {Subscriptions.Count()}");

            foreach (var subscription in Subscriptions)
            {
                // if the subscription expires in the next 2 min, renew it
                if (subscription.Value.ExpirationDateTime < DateTime.UtcNow.AddMinutes(2))
                {
                    RenewSubscription(subscription.Value);
                }
            }
        }

        private async void RenewSubscription(Subscription subscription)
        {
            Console.WriteLine($"Current subscription: {subscription.Id}, Expiration: {subscription.ExpirationDateTime}");

            var newSubscription = new Subscription
            {
                ExpirationDateTime = DateTime.UtcNow.AddMinutes(5)
            };

            await GetGraphClient()
                .Subscriptions[subscription.Id]
                .Request()
                .UpdateAsync(newSubscription);

            subscription.ExpirationDateTime = newSubscription.ExpirationDateTime;
            Console.WriteLine($"Renewed subscription: {subscription.Id}, " +
                              $"New Expiration: {subscription.ExpirationDateTime}");
        }
        [HttpGet]
        public async Task<ActionResult<string>> Get()
        {
            var newSubscription = await GetGraphClient()
                .Subscriptions
                .Request()
                .AddAsync(
                    new Microsoft.Graph.Subscription()
                    {
                        ChangeType = "updated",
                        NotificationUrl = config.Ngrok + "/api/notifications",
                        Resource = "/users",
                        ExpirationDateTime = DateTime.UtcNow.AddMinutes(5),
                        ClientState = "SecretClientState"
                    });

            Subscriptions[newSubscription.Id] = newSubscription;
            subscriptionTimer = subscriptionTimer ?? new Timer(CheckSubscriptions, null, 5000, 15000);

            return $"Subscribed. Id: {newSubscription.Id}, Expiration: {newSubscription.ExpirationDateTime}";
        }

        public async Task<ActionResult<string>> Post([FromQuery]string validationToken = null)
        {
            // handle validation
            if (!string.IsNullOrEmpty(validationToken))
            {
                Console.WriteLine($"Received Token: '{validationToken}'");
                return Ok(validationToken);
            }

            // handle notifications
            using (StreamReader reader = new StreamReader(Request.Body))
            {
                string content = await reader.ReadToEndAsync();
                Console.WriteLine(content);
                var notifications = JsonConvert.DeserializeObject<Notifications>(content);
                notifications.Items.ToList().ForEach(notification =>
                    Console.WriteLine($"Received notification: '{notification.Resource}', {notification.ResourceData?.Id}")
                );
            }
            await CheckForUpdates();  // use deltaquery to query for all updates

            return Ok();
        }

        private GraphServiceClient GetGraphClient()
        {
            return new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                var accessToken = GetAccessToken().Result;  // get an access token for Graph
                requestMessage
                    .Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                return Task.FromResult(0);
            }));
        }

        private async Task<string> GetAccessToken()
        {
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
            var result = await ConfidentialClientApplicationBuilder.Create(config.AppId)
                .WithClientSecret(config.AppSecret)
                .WithAuthority($"https://login.microsoftonline.com/{config.TenantId}")
                .WithRedirectUri("https://daemon")
                .Build()
                .AcquireTokenForClient(scopes)
                .ExecuteAsync();

            return result.AccessToken;
        }

    }
}