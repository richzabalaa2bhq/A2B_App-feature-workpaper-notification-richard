using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class SkypeBot
    {

        private readonly IConfiguration _config;
        public SkypeBot(IConfiguration config)
        {
            _config = config;
        }

        public Task<bool> SendSkypeNotif(Skype skype)
        {
            try
            {
                string msAppId = _config.GetSection("MicrosoftAppId").Value;
                string msAppPassword = _config.GetSection("MicrosoftAppPassword").Value;

                var connector = new ConnectorClient(new Uri("https://skype.botframework.com"), msAppId, msAppPassword);
                var conversation = new ConversationAccount(true, id: skype.Address);

                string message = string.Empty;
    
                List<Entity> listEntity = null;
                if (skype.Mention != null)
                {
                    var entity = new Microsoft.Bot.Schema.Mention()
                    {
                        Type = "mention",
                        Mentioned = new ChannelAccount { Id = skype.Mention.Id, Name = skype.Mention.Name },
                        Text = $"<at>{skype.Mention.Name}</at>",
                    };

                    listEntity = new List<Entity>() { entity };
                    message = $"{skype.Message} <br /> <at>{skype.Mention.Name}</at>";
                }
                else
                {
                    message = $"{skype.Message}";
                }

                var activity = new Microsoft.Bot.Schema.Activity();

                activity.Text = message;
                if(listEntity != null)
                    activity.Entities = listEntity;
                // Note on ChannelId:
                // The Bot Framework channel is identified by a unique ID.
                // For example, "skype" is a common channel to represent the Skype service.
                // We are inventing a new channel here.
                activity.ServiceUrl = "https://smba.trafficmanager.net/apis/";
                activity.ChannelId = "skype";
                //From = new ChannelAccount(id: "user", name: "Levin"),
                //Recipient = new ChannelAccount(id: "bot", name: "Bot"),
                activity.Conversation = conversation;
                activity.Timestamp = DateTime.UtcNow;
                activity.Id = Guid.NewGuid().ToString();
                activity.Type = ActivityTypes.Message;
                activity.Locale = "en-En";
                activity.TextFormat = "markdown";


                connector.Conversations.SendToConversation(activity);

                ////var members = await connector.Conversations.GetActivityMembersAsync(conversation.Id, activity.Id);
                //var members = await connector.Conversations.GetConversationMembersAsync(conversation.Id);

                //// Concatenate information about all the members into a string
                //WriteLog write = new WriteLog();
                //var sb = new StringBuilder();
                //foreach (var member in members)
                //{
                //    write.Display(member);
                //    Debug.WriteLine("--------------------------------------------");
                //}

                return Task.FromResult(true);
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSendSkypeNotif");
                return Task.FromResult(false);
            }
        }

        public Task<bool> SendSkypeImageNotif(Skype skype, string newFilename)
        {
            try
            {
                string msAppId = _config.GetSection("MicrosoftAppId").Value;
                string msAppPassword = _config.GetSection("MicrosoftAppPassword").Value;

                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "upload", "image", newFilename);

                var connector = new ConnectorClient(new Uri("https://skype.botframework.com"), msAppId, msAppPassword);
                var conversation = new ConversationAccount(true, id: skype.Address);

                string message = string.Empty;

                //Attachment attachment = new Attachment();
                //attachment = GetInlineAttachment(newFilename);

                var activity = new Microsoft.Bot.Schema.Activity();

                //activity.Text = message;
                activity.ServiceUrl = "https://smba.trafficmanager.net/apis/";
                activity.ChannelId = "skype";
                activity.Conversation = conversation;
                activity.Timestamp = DateTime.UtcNow;
                activity.Id = Guid.NewGuid().ToString();
                activity.Type = ActivityTypes.Message;
                activity.Locale = "en-En";
                activity.TextFormat = "markdown";
                activity.Attachments =  new List<Attachment>() { GetInlineAttachment(newFilename) }; ;

                connector.Conversations.SendToConversation(activity);

                return Task.FromResult(true);
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSendSkypeNotif");
                return Task.FromResult(false);
            }
        }

        private static Attachment GetInlineAttachment(string newFilename)
        {
            string startupPath = Directory.GetCurrentDirectory();
            string path = Path.Combine(startupPath, "include", "upload", "image", newFilename);

            var imageData = Convert.ToBase64String(File.ReadAllBytes(path));
            return new Attachment
            {
                Name = $"{newFilename}",
                ContentType = "image/png",
                ContentUrl = $"data:image/png;base64,{imageData}"
            };
        }

    
    }


}
