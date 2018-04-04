using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace Sonoko
{
    class Program
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";
        private static readonly DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        private static DateTime ConvertEpochToDateTime(long? epochInMiliSeconds)
        {
            return epoch.AddMilliseconds((double)epochInMiliSeconds);
        }

        private static string EmailFrom(List<MessagePartHeader> headers)
        {
            return headers.FirstOrDefault(h => h.Name.Equals("From")).Value;
        }

        static void Main(string[] args)
        {
            UserCredential credential;

            using (FileStream stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                                                                         Scopes,
                                                                         "user",
                                                                         CancellationToken.None,
                                                                         new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            int timesLooped = 1;
            while (true)
            {
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");

                IList<Message> messages = request.Execute().Messages;

                List<Message> retrievedMessages = new List<Message>();
                foreach (Message message in messages)
                {
                    try
                    {
                        retrievedMessages.Add(service.Users.Messages.Get("me", message.Id).Execute());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Issues with message id: message.Id");
                    }
                }

                foreach (Message message in retrievedMessages)
                {
                    List<MessagePartHeader> headers = message.Payload.Headers.ToList();
                    string messageFrom = EmailFrom(headers);
                    DateTime messageDate = ConvertEpochToDateTime(message.InternalDate);
                    if (messageFrom.Contains("hansen.brendan@myldsmail.net") && messageDate > DateTime.Now.AddDays(-6))
                    {
                        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"Kaigan-Dori.wav");
                        player.Play();
                    }
                }

                Console.WriteLine("No email from Brendan Found this time");
                System.Threading.Thread.Sleep(60000);
                Console.WriteLine("Repeated {0} times!", timesLooped);
                ++timesLooped;
            }
        }
    }
}
