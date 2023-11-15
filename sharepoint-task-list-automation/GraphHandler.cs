using Newtonsoft.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using System.ComponentModel.Design;
public class GraphHandler
{   
    private readonly string _siteName;
    public GraphServiceClient GraphClient { get; private set; }

    public GraphHandler(string tenantId, string clientId, string clientSecret , string siteName)
    {
        GraphClient = CreateGraphClient(tenantId, clientId, clientSecret);
        _siteName = siteName;
    }
    public GraphServiceClient CreateGraphClient(string tenantId, string clientId, string clientSecret)
    {
        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        var clientSecretCredential = new ClientSecretCredential(
            tenantId, clientId, clientSecret, options);
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        return new GraphServiceClient(clientSecretCredential, scopes);
    }
    
    private async Task<string?> GetSharePointSiteId()
    {
        var sites = await GraphClient.Sites.GetAllSites.GetAsync();
        string? siteId = sites.Value.FirstOrDefault(name => name.DisplayName == _siteName).Id;
        return siteId;
    }

    private async Task<string?> GetSharePointListId()
    {
        var siteId = await GetSharePointSiteId();
        var lists = await GraphClient.Sites[siteId].Lists.GetAsync();
        var listId = lists.Value.FirstOrDefault(name => name.DisplayName == "Tasks").Id;
        return listId;
    }

    private async Task<string?> GetSharePointListcolumnDefinitionId()
    {
        var siteId = await GetSharePointSiteId();
        var listId = await GetSharePointListId();
        var columns = await GraphClient.Sites[siteId].Lists[listId].Columns.GetAsync();
        var choicesId = columns.Value.FirstOrDefault(column => column.DisplayName == "Подрядчик").Id;
            //.Select(choice => choice.Choice.Choices);
        return choicesId;
        
    }

    public async Task PatchSharePointListColumnDefinition()
    {
        var siteId = await GetSharePointSiteId();
        var listId = await GetSharePointListId();
        var columnDefId = await GetSharePointListcolumnDefinitionId();
        
        var currentColumn = await GraphClient.Sites[siteId].Lists[listId].Columns.GetAsync();
        
        var body = new ColumnDefinition
        {
            Choice = new ChoiceColumn
            {
                Choices = new List<string> { "Choice1", "Choice2" } 
            }
            
        };
        var result = await GraphClient.Sites[siteId].Lists[listId].Columns[columnDefId].PatchAsync(body);
    }
    
    
}