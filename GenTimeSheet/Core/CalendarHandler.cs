using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

internal class CalendarHandler
{
    internal async Task<List<int>> GetMonthHolidays(int _month)
    {
        string[]? _response = await Web.GetResponse();

        int i = 0;

        while (!_response[i].Contains("blockquote"))
        {
            i++;
        }

        var holidaysString = new List<string>();

        while (!_response[i].Contains("/blockquote"))
        {
            if (_response[i].Contains("<p>"))
            {
                holidaysString.Add(_response[i]);
            }

            i++;
        }

        string monthGenitiveNames = DateTimeFormatInfo.CurrentInfo.MonthGenitiveNames[_month];

        var dates = new List<int>();

        foreach (var paragraph in holidaysString)
        {
            if (paragraph.Contains(monthGenitiveNames))
            {
                string[] s = paragraph.Replace("<p>", "").Replace(",", "").Split(" ");

                foreach (var item in s)
                {
                    if (int.TryParse(item, out int result))
                    {
                        dates.Add(result);
                    }
                }
            }
        }

        return dates;
    }

    internal Dictionary<string, List<int>> GetHolidays()
    {
        var holidays = new Dictionary<string, List<int>>();

        var calendarHandler = new CalendarHandler();

        for (int j = 0; j < 12; j++)
        {
           /* List<int> dates = calendarHandler.GetMonthHolidays(j);

            if (dates.Count > 0)
            {
                string month = DateTimeFormatInfo.CurrentInfo.MonthNames[j];

                if (holidays.ContainsKey(month))
                {
                    holidays[month].AddRange(dates);
                }
                else
                {
                    holidays.Add(month, dates);
                }
            }*/
        }

        return holidays;
    }

    internal async Task<Dictionary<string, string[]>> GetWeekends()
    {
        string[]? _response = await Web.GetResponse();

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

    internal List<int> GetMonthWeekends(int monthIndex)
    {
        string[]? _response = Web.Response;
        string[] dates = [];

        for (int i = 0; i < _response?.Length; i++)
        {
            string month = DateTimeFormatInfo.CurrentInfo.MonthNames[monthIndex];

            if (_response[i].Contains(month, StringComparison.OrdinalIgnoreCase))
            {
                int datesHTMLStringIndex = i + 14;

                dates = _response[datesHTMLStringIndex].Split("weekend\">");

                for (int j = 1; j < dates.Length; j++)
                {
                    dates[j] = dates[j].Remove(dates[j].IndexOf('<'));
                }
            }
        }

        return dates[1..].Select(int.Parse).ToList();
    }
}