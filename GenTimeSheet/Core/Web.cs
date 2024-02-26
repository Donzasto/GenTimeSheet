using System;
using System.Diagnostics;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

internal static class Web
{
    private static readonly HttpClient sharedClient = new()
    {
        BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
    };

    internal static async Task<string[]> GetResponse()
    {
        string[] responseStrings = [];

        try
        {
            using HttpResponseMessage response = await sharedClient.GetAsync("2024/");

            var stringResponse = await response.Content.ReadAsStringAsync();

            responseStrings = stringResponse.Split('\n');

        }
        catch (Exception e)
        {
            throw;
        }      
      
        return responseStrings;
    }
}