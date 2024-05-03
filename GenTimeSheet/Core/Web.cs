using System;
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

            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception(response.RequestMessage.RequestUri + " " + response.StatusCode);
            }

            var stringResponse = await response.Content.ReadAsStringAsync();

            responseStrings = stringResponse.Split('\n');
        }
        catch
        {
            throw;
        }      
      
        return responseStrings;
    }
}