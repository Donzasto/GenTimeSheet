using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GenTimeSheet.Core;

public class CalendarHandler
{
    internal async Task<List<int>> GetMonthHolidaysDates(int monthIndex)
    {
        Dictionary<string, List<int>> monthHolidaysWithNames = await GetMonthHolidays(monthIndex);

        List<int> dates = monthHolidaysWithNames.SelectMany(s => s.Value).ToList();

        return dates;
    }

    internal async Task<Dictionary<string, List<int>>> GetMonthHolidays(int monthIndex)
    {
        List<string> holidaysParagraphs;

        try
        {
            holidaysParagraphs = await GetHolidaysParagraphs();
        }
        catch
        {
            throw;
        }        

        string monthGenitiveNames = DateTimeFormatInfo.CurrentInfo.MonthGenitiveNames[monthIndex];

        var paragraphs = holidaysParagraphs.Where(p => p.Contains(monthGenitiveNames));

        // get a substring bettween '; ' and ';' or '.'
        var regex = new Regex("(?<=; )(.*?)(?=;|\\.)");

        var names = paragraphs.Select(p => regex.Match(p).Value).ToList();
        
        var holidays = new Dictionary<string, List<int>>();

        foreach (var paragraph in holidaysParagraphs)
        {
           var name = regex.Match(paragraph).Value;

            if (paragraph.Contains(monthGenitiveNames) && paragraph.Contains(name))
            {
                var dates = new List<int>();

                string[] s = paragraph.Replace("<p>", "").Replace(",", "").Split(" ");

                foreach (var item in s)
                {
                    if (int.TryParse(item, out int result))
                    {
                        dates.Add(result);
                    }
                }
                
                holidays.Add(name, dates);
            }
        }

        return holidays;
    }

    private async Task<List<string>> GetHolidaysParagraphs()
    {
        string[] _response;

        try
        {
            _response = await Web.GetResponse();
        }
        catch
        {
            throw;
        }    
        
        int i = 0;

        while (!_response[i].Contains("blockquote"))
        {
            i++;
        }

        var holidaysParagraphs = new List<string>();

        while (!_response[i].Contains("/blockquote"))
        {
            if (_response[i].Contains("<p>"))
            {
                holidaysParagraphs.Add(_response[i]);
            }

            i++;
        }

        return holidaysParagraphs;
    }

    internal async Task<List<int>> GetMonthWeekends(int monthIndex)
    {
        string[] _response;

        try
        {
            _response = await Web.GetResponse();
        }
        catch
        {
            throw;
        }

        string[] dates = [];

        for (int i = 0; i < _response?.Length; i++)
        {
            string month = DateTimeFormatInfo.CurrentInfo.MonthNames[monthIndex];

            if (_response[i].Contains($"class=\"month\">{month}", StringComparison.OrdinalIgnoreCase))
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

    /* internal async Task<List<string>> GetHolidaysNames(int monthIndex)
     {
         List<string> holidaysParagraphs = await GetHolidaysParagraphs();

         string monthGenitiveNames = DateTimeFormatInfo.CurrentInfo.MonthGenitiveNames[monthIndex];

         var paragraphs = holidaysParagraphs.Where(p => p.Contains(monthGenitiveNames));

         // get a substring bettween '; ' and ';' or '.'
         var regex = new Regex("(?<=; )(.*?)(?=;|\\.)");

         return paragraphs.Select(p => regex.Match(p).Value).ToList();
     }*/

    /*internal Dictionary<string, List<int>> GetHolidays()
    {
        var holidays = new Dictionary<string, List<int>>();

        var calendarHandler = new CalendarHandler();

        for (int j = 0; j < 12; j++)
        {
             List<int> dates = calendarHandler.GetMonthHolidays(j);

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
             }
        }

        return holidays;
    }*/

    /*  internal async Task<Dictionary<string, string[]>> GetWeekends()
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
      }*/
}