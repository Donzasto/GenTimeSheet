using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

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

            for (int i = 1; i <= 12; i++)
            {
                PopulateDays(i);
            }
        }

        private void PopulateDays(int monthNumber)
        {
            int daysColumn = 7;
            int daysRows = 5;

            var textblocks = new List<TextBlock>();

            int days = DateTime.DaysInMonth(2024, monthNumber);
            int count = 1;

            for (int j = 0; j < daysRows; j++)
            {
                for (int i = 0; i < daysColumn; i++)
                {
                    var textBlock = new TextBlock() { Text = count.ToString() };
                    count++;
                    textBlock.SetValue(Grid.ColumnProperty, i);
                    textBlock.SetValue(Grid.RowProperty, j);
                    textblocks.Add(textBlock);

                    if (count > days)
                    {
                        break;
                    }
                }
            }

            DateTimeFormatInfo dtfi = CultureInfo.CreateSpecificCulture("en-US").DateTimeFormat;

            string m = dtfi.MonthNames[monthNumber - 1];

            var grid = AnnualCalendarGrid.FindControl<Grid>(m.ToString());
            grid?.Children.AddRange(textblocks);
        }
    }
}
