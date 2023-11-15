using Microsoft.Extensions.Configuration;


var config = new ConfigurationBuilder()
    .AddJsonFile(Path.Combine(Directory.GetCurrentDirectory(),"appsettings.json"))
    .Build();
var appId = config["azureAppId"];
var siteName = config["siteName"];
var tenatId = config["tenatId"];
var clientSecret = config["clientSecret"];


var graph = new GraphHandler(tenatId, appId, clientSecret, siteName);
await graph.PatchSharePointListColumnDefinition();