using System.Collections.Generic;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private List<string>? _holidays;

    [ObservableProperty]
    private bool _enabledStart;

    [ObservableProperty]
    private List<string> _validationErrors;

    public Task Initialization { get; private set; }

    public MainViewModel()
    {
        Initialization = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        Holidays = await Web.GetHolidays();
    }

    partial void OnHolidaysChanged(List<string>? value)
    {
        if (value != null)
        {
            EnabledStart = true;
        }
    }

    public void ClickStart()
    {
        var  validation = new Validation(_holidays);

        validation.ValidateDocx();

        ValidationErrors = validation.ValidationErrors;

        var generator = new Generator(validation);

        generator.UpdateCells();
    }
}