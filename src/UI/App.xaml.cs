using System;
using System.Configuration;
using System.Data;
using System.Windows;
using overlay_gpt.Services;
using System.Text.Json;

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
        
        // 웹소켓 서비스 초기화
        InitializeWebSocket();
        
        app.Run();
    }

    private static async void InitializeWebSocket()
    {
        try
        {
            // 웹소켓 매니저 초기화
            WebSocketManager.Initialize("ws://localhost:8080/ws");
            
            // 웹소켓 서비스 시작
            await WebSocketManager.Instance.StartAsync();
            
            // 필요한 명령어 핸들러 등록
            WebSocketManager.Instance.RegisterCommandHandler("chat", async (parameters) =>
            {
                var message = parameters.GetProperty("message").GetString();
                Console.WriteLine($"채팅 메시지: {message}");
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"웹소켓 서비스 초기화 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        base.OnExit(e);
        
        // 애플리케이션 종료 시 웹소켓 서비스 정리
        WebSocketManager.Instance.StopAsync().Wait();
    }
}

