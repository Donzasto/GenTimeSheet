using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

internal static class Web
{
    internal static string[]? Response { get; private set; }

    private static readonly HttpClient sharedClient = new()
    {
        BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
    };

    static Web()
    {
        GetAsyncResponse();
    }

    internal static async void  GetAsyncResponse()
    {
        Response = await GetResponse();
    }

    internal static async Task<string[]> GetResponse()
    {
        using HttpResponseMessage response = await sharedClient.GetAsync("2024/");

        var stringResponse = await response.Content.ReadAsStringAsync();

        Response = stringResponse.Split('\n');

        return Response;
    }    
}