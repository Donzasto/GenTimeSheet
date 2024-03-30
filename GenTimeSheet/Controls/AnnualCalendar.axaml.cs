using Avalonia.Controls;
using Avalonia.Controls.Primitives;
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

        protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
        {
            PopulateAnnualCalendar();
        }

        private void PopulateAnnualCalendar()
        {
            for (int i = 0; i < 12; i++)
            {
                PopulateMonth(i);
            }
        }

        private async void PopulateMonth(int monthIndex)
        {
            var calendarHandler = new CalendarHandler();

            Dictionary<string, List<int>> holidays = null;
            List<int> weekends = null;

            try
            {
                holidays = await calendarHandler.GetMonthHolidays(monthIndex);
                weekends = await calendarHandler.GetMonthWeekends(monthIndex);
            }
            catch (Exception ex)
            {
                if (DataContext is MainViewModel mainViewModel)
                {
                    if (mainViewModel.Messages != null && !mainViewModel.Messages.Contains(ex.Message))
                    {
                        mainViewModel.Messages.Add(ex.Message);
                    }
                }
            }

            int monthNumber = monthIndex + 1;

            int daysColumn = 7;
            int daysRows = 7;

            var textblocks = new List<TextBlock>();

            string[] dayNames = { "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс" };

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

                    var name = holidays.Where(s => s.Value.Contains(dayNumber)).FirstOrDefault();

                    if (name.Value != null && name.Value.Contains(dayNumber))
                    {
                        textBlock.Classes.Add("holiday");

                        SetToolTip(textBlock, name.Key);
                    }
                    else if (weekends.Contains(dayNumber))
                    {
                        textBlock.Classes.Add("weekend");

                        SetToolTip(textBlock, "Выходной");
                    }

                    textblocks.Add(textBlock);

                    dayNumber++;
                }

                firstDayMonth = 0;
            }

            DateTimeFormatInfo dtfi = CultureInfo.CreateSpecificCulture("en-US").DateTimeFormat;

            string monthName = dtfi.MonthNames[monthIndex];

            var grid = AnnualCalendarGrid.FindControl<Grid>(monthName);

            grid?.Children.AddRange(textblocks);
        }

        private void SetToolTip(Control control, string tip)
        {
            ToolTip.SetTip(control, tip);
            ToolTip.SetShowDelay(control, 0);
        }
    }
}
