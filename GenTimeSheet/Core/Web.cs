using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

internal static class Web
{
    private static string[]? _response;

    private static readonly HttpClient sharedClient = new()
    {
        BaseAddress = new Uri("https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"),
    };

    private static async Task<string[]> GetResponse()
    {
        using HttpResponseMessage response = await sharedClient.GetAsync("2024/");

        var stringResponse = await response.Content.ReadAsStringAsync();

        _response = stringResponse.Split('\n');

        return _response;
    }

    internal static async Task<List<string>> GetHolidays()
    {
        _response = await GetResponse();

        int i = 0;

        while (!_response[i].Contains("blockquote"))
        {
            i++;
        }

        var holidays = new List<string>();

        while (!_response[i].Contains("/blockquote"))
        {
            if (_response[i].Contains("<p>"))
            {
                holidays.Add(_response[i]);
            }

            i++;
        }

        return holidays;
    }

    internal async static Task<Dictionary<string, string[]>> GetWeekends()
    {
        _response = await GetResponse();

        var weekends = new Dictionary<string, string[]>();

        int monthIndex = 0;

        for (int i = 0; i < _response.Length && monthIndex < 12; i++)
        {
            string month = DateTimeFormatInfo.CurrentInfo.MonthNames[monthIndex];

            if (_response[i].Contains(month, StringComparison.OrdinalIgnoreCase))
            {
                int datesHTMLStringIndex = i + 14;

                string[] dates = _response[datesHTMLStringIndex].Split("weekend\">");

                for (int j = 1; j < dates.Length; j++)
                {
                    dates[j] = dates[j].Remove(dates[j].IndexOf('<'));
                }

                weekends.Add(month, dates);

                monthIndex++;
            }
        }

        return weekends;
    }
}