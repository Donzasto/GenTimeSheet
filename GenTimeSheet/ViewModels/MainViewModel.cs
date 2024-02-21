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
    private string[] _response;

    [ObservableProperty]
    private Task _initialization;

    public MainViewModel()
    {
        Initialization = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        Response = await Web.GetResponse();
    }

    public async void ClickStart()
    {
        var validation = new Validation();

        await validation.ValidateDocx();

        ValidationErrors = validation.ValidationErrors;

        var generator = new Generator(validation);

        generator.UpdateCells();
    }
}