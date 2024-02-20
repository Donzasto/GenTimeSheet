using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private List<string>? _holidays;

    [ObservableProperty]
    private bool _isEnabledStart;

    [ObservableProperty]
    private List<string> _validationErrors;

    [ObservableProperty]
    private Dictionary<string, string[]> _weekends;

    public Task Initialization { get; private set; }

    public MainViewModel()
    {
        Initialization = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        Holidays = await Web.GetHolidays();
        Weekends = await Web.GetWeekends();
    }

    partial void OnHolidaysChanged(List<string>? value)
    {
        if (value != null)
        {
            IsEnabledStart = true;
        }
    }

    public void ClickStart()
    {
        var validation = new Validation(_holidays);

        validation.ValidateDocx();

        ValidationErrors = validation.ValidationErrors;

        var generator = new Generator(validation);

        generator.UpdateCells();
    }
}