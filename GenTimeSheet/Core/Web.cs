using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet
{
    internal class Web
    {
        private static string[]? _strings;

        private static readonly HttpClient sharedClient = new()
        {
            BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
        };

        internal static async Task<List<string>> GetHolidays()
        {
           await Task.Delay(5000);

            using HttpResponseMessage response = await sharedClient.GetAsync("2024/");

            var stringResponse = await response.Content.ReadAsStringAsync();

            _strings = stringResponse.Split('\n');

            int i = 0;

            while (!_strings[i].Contains("blockquote"))
            {
                i++;
            }

            var holidays = new List<string>();

            while (!_strings[i].Contains("/blockquote"))
            {
                if (_strings[i].Contains("<p>"))
                {
                    holidays.Add(_strings[i]);
                }

                i++;
            }

            return holidays;
        }
    }
}
