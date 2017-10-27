using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using BotAuth.Models;
using System.Configuration;
using BotAuth.Dialogs;
using BotAuth.AADv2;
using System.Threading;
using System.Net.Http;
using BotAuth;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using ByteSizeLib;
using MsftGraphBotQuickStartLUIS.Helpers;
using MsftGraphBotQuickStartLUIS.Extensions;

namespace MsftGraphBotQuickStart.Dialogs
{
    [LuisModel("", "", domain: "")]
    [Serializable]
    public class RootDialog : LuisDialog<IMessageActivity>
    {
        private AuthenticationOptions authenticationOptions = new AuthenticationOptions()
        {
            Authority = ConfigurationManager.AppSettings["aad:Authority"],
            ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
            ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
            Scopes = new string[] { "Files.ReadWrite" },
            RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
        };

        /// <summary>
        /// Catching the "None" intent of LUIS
        /// </summary>
        /// <param name="context"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            // No intent was found, so we're instructing the user how to use the Bot
            await context.PostAsync("I didn't understand your query...I'm just a simple bot that searches OneDrive. Try a query similar to these:<br/>'find all music'<br/>'find all .pptx files'<br/>'search for mydocument.docx'");
        }

        /// <summary>
        /// Catching the "Search Files" intent of LUIS
        /// </summary>
        /// <param name="context"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        [LuisIntent("SearchFiles")]
        public async Task SearchFiles(IDialogContext context, LuisResult result)
        {
            // Making sure there is at least one Entity that was found by LUIS
            if (result.Entities.Count == 0)
            {
                await None(context, result);
            }
            else
            {
                // Using the OneDrive search API through the Microsoft Graph
                var query = "https://graph.microsoft.com/v1.0/me/drive/search(q='{0}')?$select=id,name,size,webUrl&$top=5&$expand=thumbnails";

                // Depending on the type of Entity that is found by LUIS, we'll build the query
                if (result.Entities[0].Type == "FileName")
                {
                    // Builds a query to perform a search for a specific File Name
                    query = QueryBuilder.GetFileNameQuery(query, result.Entities[0].Entity);
                }
                else if (result.Entities[0].Type == "FileType")
                {
                    // Builds a query to perform a search based on filetype
                    query = QueryBuilder.GetFileTypeQuery(query, result.Entities[0].Entity);
                }

                // Save the query so we can run it after authenticating
                context.ConversationData.SetValue<string>("GraphQuery", query);

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph and then execute the specific action
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), this.GetOneDriveFiles, context.Activity, CancellationToken.None);
            }
        }

        /// <summary>
        /// Catching the "Delete File" intent of LUIS
        /// </summary>
        /// <param name="context"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        [LuisIntent("DeleteFile")]
        public async Task DeleteFile(IDialogContext context, LuisResult result)
        {
            // Making sure there is at least one Entity that was found by LUIS
            if (result.Entities.Count == 0)
            {
                await None(context, result);
            }
            else
            {
                // Save the retrieved FileId from the LUIS Entity so we can run it after authenticating
                context.ConversationData.SetValue<string>("FileId", result.Entities[0].Entity);

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph and then execute the specific action
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), this.DeleteSelectedFile, context.Activity, CancellationToken.None);
            }
        }

        /// <summary>
        /// Gets all the OneDrive through Microsoft Graph OneDrive Search API
        /// </summary>
        /// <param name="context">The context provvided by the Bot Framework</param>
        /// <param name="authResult">The Authentication Result provided by the Microsoft Graph</param>
        /// <returns></returns>
        private async Task GetOneDriveFiles(IDialogContext context, IAwaitable<AuthResult> authResult)
        {
            // Getting the token from the Microsoft Graph
            var tokenInfo = await authResult;

            // Get the Documents from the OneDrive of the Signed-In User
            var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, context.ConversationData.GetValue<string>("GraphQuery"));
            var items = (JArray)json.SelectToken("value");
            var reply = ((Activity)context.Activity).CreateReply();
            foreach (var item in items)
            {
                HeroCard card = new HeroCard
                {
                    Title = item.Value<string>("name"),
                    Subtitle = $"Size: {ByteSize.FromBytes(Double.Parse(item.Value<int>("size").ToString()))}",
                    Images = new List<CardImage> { new CardImage(item["thumbnails"][0]["large"].Value<string>("url")) },
                    Buttons = new List<CardAction> {
                                new CardAction(ActionTypes.OpenUrl, "View File", value: item.Value<string>("webUrl")),
                                new CardAction(ActionTypes.ImBack, "Delete", value: $"Delete file with id {item.Value<string>("id")}")
                            }
                };

                reply.Attachments.Add(card.ToAttachment());
            }

            // Build the Card as a Carousel
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            ConnectorClient client = new ConnectorClient(new Uri(context.Activity.ServiceUrl));

            // Return the HeroCard Carousel
            await client.Conversations.ReplyToActivityAsync(reply);
        }


        /// <summary>
        /// Gets all the OneDrive through Microsoft Graph OneDrive Search API
        /// </summary>
        /// <param name="context">The context provvided by the Bot Framework</param>
        /// <param name="authResult">The Authentication Result provided by the Microsoft Graph</param>
        /// <returns></returns>
        private async Task DeleteSelectedFile(IDialogContext context, IAwaitable<AuthResult> authResult)
        {
            var queryFormat = "https://graph.microsoft.com/v1.0/me/drive/items/{0}";

            // Getting the token from the Microsoft Graph
            var tokenInfo = await authResult;

            // Delete the specified Document from the OneDrive of the Signed-In User
            var json = await new HttpClient().DeleteWithAuthAsync(tokenInfo.AccessToken, String.Format(queryFormat, context.ConversationData.GetValue<string>("FileId")));

            // Send confirmation of deletion
            await context.PostAsync("File was deleted!");
        }
    }
}