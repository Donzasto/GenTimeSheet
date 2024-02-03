using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet
{
    internal class Web
    {
        private static string[]? _strings;

        private static HttpClient sharedClient = new()
        {
            BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
        };

        internal static async Task GetAsync()
        {
            using HttpResponseMessage response = await sharedClient.GetAsync("2024/");

            var stringResponse = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"{stringResponse}\n");

            _strings = stringResponse.Split('\n');
        }

        internal void GetHolidays()
        {
            for(int i = 0; i < _strings?.Length; i++)
            {
                if (_strings[i].Contains("blockquote"))
                {

                }
            }
        }
    }
}
