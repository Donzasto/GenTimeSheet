using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.ObjectModel;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private ObservableCollection<string>? _messages = [];

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
        catch (System.Exception ex)
        {
            HasRequestException = true;

            Messages?.Add(ex.Message);
        }
    }

    public async void ClickStart()
    {
        var validation = new Validation("1.docx", "2.docx");

        await validation.ValidateDocx();

        Messages = new ObservableCollection<string>(validation.ValidationErrors);

        validation.ValidationErrors.ForEach(v => Messages.Add(v));

        var generator = new Generator(validation);

        try
        {
            await generator.UpdateCells();

            Messages.Add("Файл успшено создан");
        }
        catch (System.Exception ex)
        {
            Messages.Add(ex.Message);
        }
    }
}