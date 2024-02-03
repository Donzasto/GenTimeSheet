using GenTimeSheet.Core;
using System.Diagnostics;

namespace GenTimeSheet.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private Generator _generator = new();

    public MainViewModel()
    {
        
    }

    public void ClickStart()
    {
        if (_generator != null)
        {
            Debug.WriteLine("sdf");
        }
    }
}
