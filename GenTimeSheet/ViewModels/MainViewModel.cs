using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private List<string>? _holidays;

    [ObservableProperty]
    private bool _enabledStart;

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
        /*var generator = new Generator();

        if (generator != null)
        {
            Debug.WriteLine("sdf");
        }*/

        if (Initialization != null)
        {
            Debug.WriteLine("init not null");
        }
    }
}
