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
        // If the validationToken is not null or empty, return it as ContentResult
        if (!string.IsNullOrEmpty(validationToken))
        {
            return CreateContentResult(validationToken);
        }

        // Read and parse the webhook payload from the Request Body
        JObject webhookPayload = await ReadWebhookPayload();

        // If the webhookPayload is not null, process the data
        if (webhookPayload != null)
        {
            return await ProcessWebhookData(webhookPayload);
        }

        // Return BadRequest if the validationToken and webhookPayload are both null or empty
        return BadRequest();
    }

    private ContentResult CreateContentResult(string validationToken)
    {
        return new ContentResult
        {
            Content = validationToken,
            ContentType = "text/plain",
            StatusCode = 200
        };
    }

    private async Task<JObject> ReadWebhookPayload()
    {
        using var reader = new StreamReader(Request.Body);
        var webhookPayloadStr = await reader.ReadToEndAsync();

        if (!string.IsNullOrEmpty(webhookPayloadStr))
        {
           var jsonObject = JObject.Parse(webhookPayloadStr);
           return (JObject) jsonObject["value"][0];
        }

        return null;
    }

    private async Task<IActionResult> ProcessWebhookData(JObject webhookPayload)
    {
        string resource = webhookPayload["resource"]?.ToString();
        string tenantId = webhookPayload["tenantId"]?.ToString();

        if (resource != null && tenantId != null)
        {
            string siteId, listId;

            ExtractIds(resource, out siteId, out listId);

            await SendDataToPort(siteId, listId, tenantId);
            return Ok();
        }
        
        return BadRequest();
    }

    private void ExtractIds(string resource, out string siteId, out string listId)
    {
        var siteIdEndPos = resource.IndexOf("/lists/");
        siteId = resource.Substring(0, siteIdEndPos);

        var listIdStartPos = resource.LastIndexOf("/") + 1;
        listId = resource.Substring(listIdStartPos);

        siteId = siteId.Replace("sites/", ""); // to remove "sites/" from the start
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
