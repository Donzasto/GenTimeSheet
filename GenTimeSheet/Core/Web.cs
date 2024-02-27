using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

public static class Web
{
    private static readonly HttpClient sharedClient = new()
    {
        BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
    };

    public static async Task<string[]> GetResponse()
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