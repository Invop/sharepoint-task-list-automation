using Newtonsoft.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using System.ComponentModel.Design;
public class GraphHandler
{   
    private readonly string _siteName;
    private readonly string _clientSecret;
    public GraphServiceClient GraphClient { get; private set; }

    public GraphHandler(string tenantId, string clientId, string clientSecret , string siteName)
    {
        GraphClient = CreateGraphClient(tenantId, clientId, clientSecret);
        _siteName = siteName;
        _clientSecret = clientSecret;
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

    private async Task<string?> GetSharePointListId(bool isTaskList)
    {   
        string displayName;
        if (isTaskList)
        {
            displayName = "Tasks";
        }
        else
        {
            displayName = "Подрядчик";
        }
        
        var siteId = await GetSharePointSiteId();
        var lists = await GraphClient.Sites[siteId].Lists.GetAsync();
        var listId = lists.Value.FirstOrDefault(name => name.DisplayName == displayName).Id;
        return listId;
    }

    private async Task<string?> GetSharePointListСolumnDefinitionId()
    {
        string name = "Подрядчик";
        string? listId = await GetSharePointListId(true);
        var siteId = await GetSharePointSiteId();
        var columns = await GraphClient.Sites[siteId].Lists[listId].Columns.GetAsync();
        var choicesId = columns.Value.FirstOrDefault(column => column.DisplayName == name).Id;
            //.Select(choice => choice.Choice.Choices);
        return choicesId;
        
        
    }

    private async Task<List<string>> GetSharePointContractors()
    {
        var listId = await GetSharePointListId(false);
        var siteId = await GetSharePointSiteId();

        var response = await GraphClient.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Expand = new string[]
                { "fields($select=Title)" };
        });
        var listItems = new List<ListItem>();
        var pageIterator = PageIterator<ListItem, ListItemCollectionResponse>.CreatePageIterator(GraphClient, response,item=>
        {
            listItems.Add(item);
            return true;
        });
        
        await pageIterator.IterateAsync();

        try
        {
            List<string> titleList = listItems.Select(items =>
            {
                var name = items.Fields.AdditionalData.TryGetValue("Title", out var nameValue)
                    ? nameValue.ToString()
                    : null;
                return name;
            }).ToList();

            return titleList;
        }
        catch (Exception ex) 
        {
            Console.WriteLine(ex.Message);
        }
        
        return null;

    }

    public async Task PatchSharePointListColumnDefinition()
    {
        // Retrieve the ID of the SharePoint site
        var siteId = await GetSharePointSiteId();

        // Retrieve the ID of the task list in SharePoint.
        var listIdPrevChoices = await GetSharePointListId(true);

        // Retrieve the ID of the "Contractors (selection)" column in the SharePoint list.
        var columnDefIdPrevChoices = await GetSharePointListСolumnDefinitionId();

        // Fetch the list of contractors from SharePoint.
        List<string> newChoices = await GetSharePointContractors();

        // Get the current definition of the "Contractors (selection)" column in the SharePoint list.
        var prevChoiceDef = await GraphClient.Sites[siteId].Lists[listIdPrevChoices].Columns[columnDefIdPrevChoices].GetAsync();

        // Retrieves a list of "choices" from a column definition.
        List<string> prevChoices = prevChoiceDef.Choice.Choices;
        
        
        Console.WriteLine(newChoices.Count + " " + prevChoices.Count);
        if(newChoices.Count!=prevChoices.Count || !newChoices.SequenceEqual(prevChoices))
        {
            Console.WriteLine("changed");
            var body = new ColumnDefinition
            {
                Choice = new ChoiceColumn
                {
                    Choices = new List<string>()
                }
                
            };
            body.Choice.Choices.AddRange(newChoices); 
            await GraphClient.Sites[siteId].Lists[listIdPrevChoices].Columns[columnDefIdPrevChoices].PatchAsync(body);
        }
    }

    public async Task SubscribeToChanges()
    {   
        
        var subs = await GraphClient.Subscriptions.GetAsync();
        foreach (var pair in subs.Value)
        {
            Console.WriteLine(pair.Id);
            await GraphClient.Subscriptions[pair.Id].DeleteAsync();
        }
        
        
        var siteId = await GetSharePointSiteId();
        var listId = await GetSharePointListId(false);
        var webhookEndpoint = "https://22f1-178-212-196-200.ngrok-free.app/Notifications";
        var subscription = new Subscription
        {
            Resource = $"sites/{siteId}/lists/{listId}",
            ChangeType = "updated",
            NotificationUrl = webhookEndpoint,
            
            ExpirationDateTime = DateTime.UtcNow.AddMinutes(43200)// Adjust as needed
        };

        try
        {
            var newSubscription = await GraphClient.Subscriptions
                .PostAsync(subscription);

            Console.WriteLine($"Subscription ID: {newSubscription.Id}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error subscribing to changes: {ex.Message}");
        }
    }
}