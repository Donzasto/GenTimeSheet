using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using DocumentFormat.OpenXml.Spreadsheet;
using GenTimeSheet.Core;
using GenTimeSheet.ViewModels;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GenTimeSheet.Controls
{
    public partial class AnnualCalendar : UserControl
    {
        public AnnualCalendar()
        {
            InitializeComponent();
        }
        // binding property style classes.blue =isEnabled
        protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
        {
          /*  var calendarHandler = new CalendarHandler();

            Dictionary<string, List<int>> holidays = calendarHandler.GetHolidays();

            for (int i = 0; i < 12; i++)
            {
                string monthName = DateTimeFormatInfo.CurrentInfo.MonthNames[i];

                List<int> monthHolidays = holidays[monthName];

                PopulateDays1(i, monthHolidays);
            }*/

            for (int i = 0; i < 12; i++)
            {
                PopulateDays(i);
            }
        }

        private void PopulateDays1(int monthIndex, List<int> holidays)
        {
            int monthNumber = monthIndex + 1;

            int daysColumn = 7;
            int daysRows = 7;

            var textblocks = new List<TextBlock>();

            string[] dayNames = { "Ïí", "Âò", "Ñð", "×ò", "Ïò", "Ñá", "Âñ" };

            for (int i = 0; i < daysColumn; i++)
            {
                var textBlock = new TextBlock() { Text = dayNames[i] };
                textBlock.SetValue(Grid.ColumnProperty, i);
                textBlock.SetValue(Grid.RowProperty, 0);
                textblocks.Add(textBlock);
            }

            int daysInMonth = DateTime.DaysInMonth(2024, monthNumber);
            int dayNumber = 1;

            int firstDayMonth = (int)new DateTime(2024, monthNumber, 1).DayOfWeek;

            if (firstDayMonth == 0)
            {
                firstDayMonth = 7;
            }

            firstDayMonth--;

            for (int j = 1; j < daysRows; j++)
            {
                for (; firstDayMonth < daysColumn; firstDayMonth++)
                {
                    if (dayNumber > daysInMonth)
                    {
                        break;
                    }

                    var textBlock = new TextBlock
                    {
                        Text = dayNumber.ToString()
                    };

                    textBlock.SetValue(Grid.ColumnProperty, firstDayMonth);
                    textBlock.SetValue(Grid.RowProperty, j);

                    if (holidays.Contains(dayNumber))
                    {
                        textBlock.Classes.Add("holiday");
                    }

                    textblocks.Add(textBlock);

                    dayNumber++;
                }

                firstDayMonth = 0;
            }

            DateTimeFormatInfo dtfi = CultureInfo.CreateSpecificCulture("en-US").DateTimeFormat;

            string monthName = dtfi.MonthNames[monthIndex];

            var grid = AnnualCalendarGrid.FindControl<Grid>(monthName.ToString());
            grid?.Children.AddRange(textblocks);
        }

        private void PopulateDays(int monthIndex)
        {
            var calendarHandler = new CalendarHandler();

            List<int> holidays = calendarHandler.GetMonthHolidays(monthIndex);
            List<int> weekends = calendarHandler.GetMonthWeekends(monthIndex);

            int monthNumber = monthIndex + 1;

            int daysColumn = 7;
            int daysRows = 7;

            var textblocks = new List<TextBlock>();

            string[] dayNames = { "Ïí", "Âò", "Ñð", "×ò", "Ïò", "Ñá", "Âñ" };

            for (int i = 0; i < daysColumn; i++)
            {
                var textBlock = new TextBlock() { Text = dayNames[i] };
                textBlock.SetValue(Grid.ColumnProperty, i);
                textBlock.SetValue(Grid.RowProperty, 0);
                textblocks.Add(textBlock);
            }

            int daysInMonth = DateTime.DaysInMonth(2024, monthNumber);
            int dayNumber = 1;

            int firstDayMonth = (int)new DateTime(2024, monthNumber, 1).DayOfWeek;

            if (firstDayMonth == 0)
            {
                firstDayMonth = 7;
            }

            firstDayMonth--;

            for (int j = 1; j < daysRows; j++)
            {
                for (; firstDayMonth < daysColumn; firstDayMonth++)
                {
                    if (dayNumber > daysInMonth)
                    {
                        break;
                    }

                    var textBlock = new TextBlock
                    {
                        Text = dayNumber.ToString()
                    };

                    textBlock.SetValue(Grid.ColumnProperty, firstDayMonth);
                    textBlock.SetValue(Grid.RowProperty, j);

                    if (holidays.Contains(dayNumber))
                    {
                        textBlock.Classes.Add("holiday");
                    }
                    else if (weekends.Contains(dayNumber))
                    {
                        textBlock.Classes.Add("weekend");
                    }

                    textblocks.Add(textBlock);

                    dayNumber++;
                }

                firstDayMonth = 0;
            }

            DateTimeFormatInfo dtfi = CultureInfo.CreateSpecificCulture("en-US").DateTimeFormat;

            string monthName = dtfi.MonthNames[monthIndex];

            var grid = AnnualCalendarGrid.FindControl<Grid>(monthName.ToString());
            grid?.Children.AddRange(textblocks);
        }
    }
}
