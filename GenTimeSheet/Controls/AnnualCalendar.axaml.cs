using Avalonia.Controls;
using Avalonia.Controls.Primitives;
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
            Dictionary<string, string[]> weekends;

            if (DataContext is MainViewModel mainViewModel)
            {
                weekends = mainViewModel.Weekends;

                for (int i = 1; i <= 12; i++)
                {
                    PopulateDays(i, weekends[DateTimeFormatInfo.CurrentInfo.MonthNames[i - 1]]);
                }
            }
        }

        private void PopulateDays(int monthNumber, string[] weekend)
        {
            int daysColumn = 7;
            int daysRows = 6;

            var textblocks = new List<TextBlock>();

            int days = DateTime.DaysInMonth(2024, monthNumber);
            int count = 1;

            string[] dayNames = { "Ïí", "Âò", "Ñð", "×ò", "Ïò", "Ñá", "Âñ" };

            for (int i = 0; i < daysColumn; i++)
            {
                var textBlock = new TextBlock() { Text = dayNames[i] };
                textBlock.SetValue(Grid.ColumnProperty, i);
                textBlock.SetValue(Grid.RowProperty, 0);
                textblocks.Add(textBlock);
            }

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
                    var textBlock = new TextBlock
                    {
                        Text = count.ToString()
                    };

                    textBlock.SetValue(Grid.ColumnProperty, firstDayMonth);
                    textBlock.SetValue(Grid.RowProperty, j);

                    if (weekend.Contains(count.ToString()))
                    {
                        textBlock.Classes.Add("weekend");
                    }

                    count++;

                    textblocks.Add(textBlock);

                    if (count > days)
                    {
                        break;
                    }

                }
                firstDayMonth = 0;
            }

            DateTimeFormatInfo dtfi = CultureInfo.CreateSpecificCulture("en-US").DateTimeFormat;

            string monthName = dtfi.MonthNames[monthNumber - 1];

            var grid = AnnualCalendarGrid.FindControl<Grid>(monthName.ToString());
            grid?.Children.AddRange(textblocks);
        }
    }
}
