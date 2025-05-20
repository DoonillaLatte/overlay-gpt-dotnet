using System;
using System.Configuration;
using System.Data;
using System.Windows;
using overlay_gpt.Services;
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
        
        // LogWindow를 먼저 띄운다
        LogWindow.Instance.Show();
        
        // 웹소켓 서비스 초기화
        InitializeWebSocket();
        
        app.Run();
    }

    private static async void InitializeWebSocket()
    {
        try
        {
            // 웹소켓 매니저 초기화
            overlay_gpt.Services.WebSocketManager.Initialize("http://localhost:8080/ws/", LogWindow.Instance);
            
            // 연결 이벤트 핸들러 등록
            overlay_gpt.Services.WebSocketManager.Instance.OnConnected += () =>
            {
                LogWindow.Instance.Log("웹소켓 서버에 연결되었습니다.");
            };
            
            // 웹소켓 서비스 시작
            var wsTask = overlay_gpt.Services.WebSocketManager.Instance.StartAsync();

            // 서버가 열렸을 때 메시지 출력
            LogWindow.Instance.Log("웹소켓 서버가 정상적으로 열렸습니다.");
            
            // 필요한 명령어 핸들러 등록
            RegisterMessageHandlers();
        }
        catch (Exception ex)
        {
            // LogWindow에 실패 메시지 출력
            LogWindow.Instance.Log($"웹소켓 서비스 초기화 중 오류 발생: {ex.Message}");
            MessageBox.Show($"웹소켓 서비스 초기화 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static void RegisterMessageHandlers()
    {
        try
        {
            // 채팅 메시지 핸들러
            overlay_gpt.Services.WebSocketManager.Instance.RegisterMessageHandler("chat", async (parameters) =>
            {
                try
                {
                    if (parameters.TryGetProperty("message", out var messageElement))
                    {
                        var message = messageElement.GetString();
                        if (!string.IsNullOrEmpty(message))
                        {
                            LogWindow.Instance.Log($"채팅 메시지 수신: {message}");
                            await Task.Run(() => Console.WriteLine($"채팅 메시지: {message}"));
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogWindow.Instance.Log($"채팅 메시지 처리 중 오류 발생: {ex.Message}");
                }
            });

            // 시스템 메시지 핸들러
            overlay_gpt.Services.WebSocketManager.Instance.RegisterMessageHandler("system", (parameters) =>
            {
                try
                {
                    if (parameters.TryGetProperty("message", out var messageElement))
                    {
                        var message = messageElement.GetString();
                        if (!string.IsNullOrEmpty(message))
                        {
                            LogWindow.Instance.Log($"시스템 메시지: {message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogWindow.Instance.Log($"시스템 메시지 처리 중 오류 발생: {ex.Message}");
                }
            });
        }
        catch (Exception ex)
        {
            LogWindow.Instance.Log($"메시지 핸들러 등록 중 오류 발생: {ex.Message}");
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        base.OnExit(e);
        
        // 애플리케이션 종료 시 웹소켓 서비스 정리
        overlay_gpt.Services.WebSocketManager.Instance.StopAsync().Wait();
    }
}

