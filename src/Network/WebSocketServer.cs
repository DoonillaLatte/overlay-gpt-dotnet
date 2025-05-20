using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Runtime.InteropServices;

namespace overlay_gpt.Network;

public class WebSocketServer
{
    private IHost? _host;
    private readonly int _port;
    private bool _isRunning;

    [DllImport("kernel32.dll")]
    private static extern bool AllocConsole();

    public WebSocketServer(int port = 8080)
    {
        _port = port;
        AllocConsole(); // 콘솔 창 생성
        Console.WriteLine($"웹소켓 서버가 초기화되었습니다. (포트: {port})");
    }

    public async Task StartAsync()
    {
        if (_isRunning)
        {
            Console.WriteLine("웹소켓 서버가 이미 실행 중입니다.");
            return;
        }

        try
        {
            _host = Host.CreateDefaultBuilder()
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseUrls($"http://localhost:{_port}");
                    webBuilder.Configure(app =>
                    {
                        // CORS 설정을 UseRouting 전에 배치
                        app.UseCors(builder => builder
                            .SetIsOriginAllowed(_ => true)
                            .AllowAnyMethod()
                            .AllowAnyHeader()
                            .DisallowCredentials());

                        app.UseRouting();
                        app.UseWebSockets(new WebSocketOptions
                        {
                            KeepAliveInterval = TimeSpan.FromSeconds(120),
                            ReceiveBufferSize = 4 * 1024
                        });

                        app.UseEndpoints(endpoints =>
                        {
                            endpoints.MapHub<ChatHub>("/chatHub");
                        });
                    });
                })
                .ConfigureServices(services =>
                {
                    services.AddCors(options =>
                    {
                        options.AddDefaultPolicy(builder =>
                        {
                            builder.SetIsOriginAllowed(_ => true)
                                   .AllowAnyMethod()
                                   .AllowAnyHeader()
                                   .DisallowCredentials();
                        });
                    });
                    
                    services.AddSignalR(options =>
                    {
                        options.EnableDetailedErrors = true;
                        options.MaximumReceiveMessageSize = 102400;
                        options.KeepAliveInterval = TimeSpan.FromSeconds(30);  // 30초마다 keep-alive
                        options.ClientTimeoutInterval = TimeSpan.FromSeconds(120);  // 클라이언트 타임아웃 120초
                        options.HandshakeTimeout = TimeSpan.FromSeconds(60);  // 핸드셰이크 타임아웃 60초
                    });

                    // SocketIOConnection 서비스 등록
                    services.AddSingleton<SocketIOConnection>();
                })
                .Build();

            await _host.StartAsync();
            _isRunning = true;
            
            // Socket.IO 연결 시작
            var socketIOConnection = _host.Services.GetRequiredService<SocketIOConnection>();
            await socketIOConnection.ConnectAsync();
            
            Console.WriteLine("==========================================");
            Console.WriteLine($"웹소켓 서버가 시작되었습니다.");
            Console.WriteLine($"URL: http://localhost:{_port}/chatHub");
            Console.WriteLine($"상태: 실행 중");
            Console.WriteLine("==========================================");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"서버 시작 중 오류 발생: {ex.Message}");
            _isRunning = false;
        }
    }

    public async Task StopAsync()
    {
        Console.WriteLine("웹소켓 서버를 종료하는 중...");
        
        if (_host != null)
        {
            await _host.StopAsync();
            _host.Dispose();
        }
        
        _isRunning = false;
        
        Console.WriteLine("==========================================");
        Console.WriteLine("웹소켓 서버가 종료되었습니다.");
        Console.WriteLine("==========================================");
    }
} 