using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Text.Json;
using System.Threading.Tasks;

namespace overlay_gpt;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    [STAThread]
    public static void Main()
    {
        App app = new App();
        app.InitializeComponent();
        
        // LogWindow를 먼저 띄운다(배포용으로 숨김처리)
        //LogWindow.Instance.Show();
        
        app.Run();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        base.OnExit(e);
    }
}

