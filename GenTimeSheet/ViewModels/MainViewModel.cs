using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private Generator _generator = new();

    [ObservableProperty]
    private List<string>? _holidays;

    public Task Initialization { get; private set; }

    public MainViewModel()
    {
        Initialization = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        Holidays = await Web.GetHolidays();
    }

    public void ClickStart()
    {
        if (_generator != null)
        {
            Debug.WriteLine("sdf");
        }
    }
}
