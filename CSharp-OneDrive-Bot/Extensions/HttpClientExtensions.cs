using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace MsftGraphBotQuickStartLUIS.Extensions
{
    public static class HttpClientExtensions
    {
        public static async Task<Boolean> DeleteWithAuthAsync(this HttpClient client, string accessToken, string endpoint)
         {
             client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
             client.DefaultRequestHeaders.Add("Accept", "application/json");
             using (var response = await client.DeleteAsync(endpoint))
             {
                return response.IsSuccessStatusCode;
             }
         }
    }
}