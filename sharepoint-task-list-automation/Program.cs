using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Process = System.Diagnostics.Process;

internal class Program
{
    public static async Task Main(string[] args)
    {
        var config = new ConfigurationBuilder()
            .AddJsonFile(Path.Combine(Directory.GetCurrentDirectory(),"appsettings.json"))
            .Build();
        var appId = config["azureAppId"];
        var siteName = config["siteName"];
        var tenantId = config["tenantId"];
        var clientSecret = config["clientSecret"];

        var graph = new GraphHandler(tenantId, appId, clientSecret, siteName);
        await graph.SubscribeToChanges();
        
        await StartServer();




     
    }
    private static async Task StartServer()
    {
        HttpListener listener = new HttpListener();
        listener.Prefixes.Add("http://localhost:5002/");
        listener.Start();
        while (true)
        {   
            HttpListenerContext context = await listener.GetContextAsync();
            await HandleRequest(context);
        }
    }

    private static async Task HandleRequest(HttpListenerContext context)
    {
        string input;
        using (var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
        {
            input = await reader.ReadToEndAsync();
        }
        Console.WriteLine("Received request: " + input);
        JObject json = JObject.Parse(input);
        var siteId = json["SiteId"]?.ToString();
        var listId = json["ListId"]?.ToString();
        var tenantId = json["TenantId"]?.ToString();

        HttpListenerResponse response = context.Response;

        string responseContent = "{ \"Status\": \"Success\" }"; 

        byte[] buffer = Encoding.UTF8.GetBytes(responseContent);

        response.ContentType = "application/json";
        response.ContentLength64 = buffer.Length;
        response.StatusCode = 200;

        using (Stream output = response.OutputStream) 
        {
            await output.WriteAsync(buffer, 0, buffer.Length);
        }
    }
    
}