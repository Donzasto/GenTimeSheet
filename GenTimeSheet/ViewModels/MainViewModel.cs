using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private List<string> _validationErrors;

    [ObservableProperty]
    private string[] _weekends;

    public Task Initialization { get; private set; }

    public MainViewModel()
    {
        Initialization = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        Weekends = await Web.GetResponse();
    }

    public void ClickStart()
    {
        var validation = new Validation();

        validation.ValidateDocx();

        ValidationErrors = validation.ValidationErrors;

        var generator = new Generator(validation);

        generator.UpdateCells();
    }
}