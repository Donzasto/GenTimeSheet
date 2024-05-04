using System.Collections;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;

namespace GenTimeSheet.ViewModels;

internal partial class MainViewModel : ViewModelBase
{
    [ObservableProperty]
    private ObservableCollection<string>? _messages = [];

    [ObservableProperty]
    private string? _requestException;

    [ObservableProperty]
    private bool _hasRequestException;

    [ObservableProperty]
    private ObservableCollection<IEnumerable> _currentMonthTable;

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

    internal async void ClickStart()
    {      
        try
        {            
            var validation = new Validation("1.docx", "2.docx");

            await validation.ValidateDocx();

            var currentMonthTable = new Serialize(validation.CurrentMonthTable).GetTable();

            CurrentMonthTable = new ObservableCollection<IEnumerable>(currentMonthTable);

            Messages = new ObservableCollection<string>(validation.ValidationErrors);
            
            validation.ValidationErrors.ForEach(v => Messages.Add(v));

            var generator = new Generator(validation);

            await generator.Run();

            Messages.Add("Файл успшено создан");
        }
        catch (System.Exception ex)
        {
            Messages.Add(ex.Message);
        }
    }
}