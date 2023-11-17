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
        var tenatId = config["tenatId"];
        var clientSecret = config["clientSecret"];

        await StartServer();




     
    }
    private static async Task StartServer()
    {
        HttpListener listener = new HttpListener();
        // Добавление префикса, означающего, что сервер будет обслуживать все запросы по адресу "http://localhost:5002/"
        listener.Prefixes.Add("http://localhost:5002/");
        // Запуск сервера для прослушивания входящих запросов
        listener.Start();
        while (true)
        {   
            // Ожидание входящего запроса и получение контекста для дальнейшей обработки
            HttpListenerContext context = await listener.GetContextAsync();
            // Обработка входящего запроса
            await HandleRequest(context.Request);
        }
    }

    private static async Task HandleRequest(HttpListenerRequest request)
    {
        string input;
        // Создание потока для чтения содержимого запроса в асинхронном режиме
        using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
        {   // Чтение всего содержимого запроса и сохранение его в переменную input
            input = await reader.ReadToEndAsync();
        }
        // Вывод на консоль полученного запроса
        Console.WriteLine("Received request: " + input);
        // Преобразование строкового значения запроса в объект JSON
        JObject json = JObject.Parse(input);
        // Извлечение значения свойства "SiteId","ListId","TenantId" из объекта JSON и сохранение его в переменную siteId
        var siteId = json["SiteId"]?.ToString();
        var listId = json["ListId"]?.ToString();
        var tenatId = json["TenantId"]?.ToString();
    }
    
}