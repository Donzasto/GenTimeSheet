using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.Generic;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private List<string>? _validationErrors;

    [ObservableProperty]
    private string? _requestException;

    [ObservableProperty]
    private bool _hasRequestException;

    public MainViewModel()
    {
        GetResponse();
    }

    private async void GetResponse()
    {
        try
        {
            await Web.GetResponse();
        }
        catch (System.Exception e )
        {
            HasRequestException = true;
            RequestException = $"Ошибка синхронизации.{e.Message}";
        }
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