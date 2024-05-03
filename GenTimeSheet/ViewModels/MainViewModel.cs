using CommunityToolkit.Mvvm.ComponentModel;
using GenTimeSheet.Core;
using System.Collections.Generic;
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

    public ObservableCollection<Crockery> CrockeryList { get; set; }

    public class Crockery
    {
        public string Title { get; set; }
        public int Number { get; set; }

        public Crockery(string title, int number)
        {
            Title = title;
            Number = number;
        }
    }

    public MainViewModel()
    {
        GetResponse();

        CrockeryList = new ObservableCollection<Crockery>(new List<Crockery>
            {
                new Crockery("dinner plate", 12),
                new Crockery("side plate", 12),
                new Crockery("breakfast bowl", 6),
                new Crockery("cup", 10),
                new Crockery("saucer", 10),
                new Crockery("mug", 6),
                new Crockery("milk jug", 1)
            });
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
        try
        {            
            var validation = new Validation("1.docx", "2.docx");

            await validation.ValidateDocx();
    
            Messages = new ObservableCollection<string>(validation.ValidationErrors);

            validation.ValidationErrors.ForEach(v => Messages.Add(v));

            var generator = new Generator(validation);

            await generator.UpdateCells();

            Messages.Add("Файл успшено создан");
        }
        catch (System.Exception ex)
        {
            Messages.Add(ex.Message);
        }
    }
}