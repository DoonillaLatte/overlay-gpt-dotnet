using System;
using System.Configuration;
using System.Data;
using System.Windows;

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
        app.Run();
    }
}

