using System.Text;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace sharepoint_web_hook.Controllers;
[ApiController]
[Route("[controller]")]
public class NotificationsController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> Validate([FromQuery] string validationToken = null)
    {

        Console.WriteLine(200);
        if (!string.IsNullOrEmpty(validationToken))
        {
            return new ContentResult
            {
                Content = validationToken,
                ContentType = "text/plain",
                StatusCode = 200
            };
        }
    
        using var reader = new StreamReader(Request.Body);
        var webhookPayloadStr = await reader.ReadToEndAsync();
        JObject webhookPayload = null;
        if (!string.IsNullOrEmpty(webhookPayloadStr))
        {
            var jsonObject = JObject.Parse(webhookPayloadStr);
            webhookPayload = (JObject) jsonObject["value"][0];
        }

        if (webhookPayload != null)
        {
            var resource = webhookPayload["resource"]?.ToString();
            var tenantId = webhookPayload["tenantId"]?.ToString();

            if (resource != null && tenantId != null)
            {
                var siteIdEndPos = resource.IndexOf("/lists/");
                var siteId = resource.Substring(0, siteIdEndPos);
                
                var listIdStartPos = resource.LastIndexOf("/") + 1;
                var listId = resource.Substring(listIdStartPos);

                siteId = siteId.Replace("sites/", ""); // to remove "sites/" from the start

                await SendDataToPort(siteId, listId, tenantId);
            }

            return Ok();
        }

        return BadRequest();
    }

    private readonly HttpClient _httpClient = new HttpClient();

    private async Task SendDataToPort(string siteId, string listId, string tenantId)
    {
        var data = new
        {
            SiteId = siteId,
            ListId = listId,
            TenantId = tenantId
        };

        var json = JsonConvert.SerializeObject(data);
        var httpContent = new StringContent(json, Encoding.UTF8, "application/json");

        var httpResponse = await _httpClient.PostAsync("http://localhost:5002", httpContent);

        if (httpResponse.IsSuccessStatusCode)
        {
            Console.WriteLine("Data successfully sent");
        }
        else
        {
            Console.WriteLine($"Failed to send. Status code: {httpResponse.StatusCode}");
        }
    }
}
