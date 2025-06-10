/*
 * WithVueAsHost 클래스 사용 방법
 * 
 * 1. 서버 초기화 및 시작
 *    var server = new WithVueAsHost(port: 3000);
 *    await server.StartAsync();
 * 
 * 2. 메시지 전송 방법
 *    - 모든 클라이언트에게 전송:
 *      await server.SendMessageToAll(new { message = "안녕하세요" });
 *    
 *    - 특정 클라이언트에게 전송:
 *      await server.SendMessageToClient(connectionId, new { message = "안녕하세요" });
 *    
 *    - 특정 그룹에게 전송:
 *      await server.SendMessageToGroup(groupName, new { message = "안녕하세요" });
 * 
 * 3. 서버 종료
 *    await server.StopAsync();
 * 
 * Vue.js 클라이언트 측 연결 예시:
 * const connection = new signalR.HubConnectionBuilder()
 *     .withUrl("http://localhost:8080/vueHub")
 *     .build();
 * 
 * connection.on("ReceiveMessage", (message) => {
 *     console.log("받은 메시지:", message);
 * });
 * 
 * await connection.start();
 */

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Runtime.InteropServices;
using Microsoft.AspNetCore.SignalR;
using System.Text.Json;
using System.Diagnostics;
using Newtonsoft.Json.Linq;

namespace overlay_gpt.Network;

public class ChatHub : Hub
{
    private readonly ProcessVueMessage _messageProcessor;

    public ChatHub(ProcessVueMessage messageProcessor)
    {
        _messageProcessor = messageProcessor;
    }

    public override async Task OnConnectedAsync()
    {
        Console.WriteLine("==========================================");
        Console.WriteLine($"새로운 클라이언트가 연결되었습니다.");
        Console.WriteLine($"연결 ID: {Context.ConnectionId}");
        Console.WriteLine($"연결 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        Console.WriteLine("==========================================");
        await base.OnConnectedAsync();
    }

    public override async Task OnDisconnectedAsync(Exception? exception)
    {
        Console.WriteLine("==========================================");
        Console.WriteLine($"클라이언트 연결이 해제되었습니다.");
        Console.WriteLine($"연결 ID: {Context.ConnectionId}");
        Console.WriteLine($"연결 해제 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        if (exception != null)
        {
            Console.WriteLine($"연결 해제 사유: {exception.Message}");
        }
        Console.WriteLine("==========================================");
        await base.OnDisconnectedAsync(exception);
    }

    public async Task SendMessage(object message)
    {
        try
        {
            JObject jObject;
            if (message is string jsonString)
            {
                jObject = JObject.Parse(jsonString);
            }
            else if (message is JsonElement jsonElement)
            {
                jObject = JObject.Parse(jsonElement.GetRawText());
            }
            else
            {
                jObject = JObject.FromObject(message);
            }

            var command = jObject["command"]?.ToString();
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 명령어: {command}");

            if (string.IsNullOrEmpty(command))
            {
                throw new Exception("명령어가 지정되지 않았습니다.");
            }

            await _messageProcessor.ProcessMessage(Context.ConnectionId, jObject);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            var response = new
            {
                status = "error",
                message = $"메시지 처리 중 오류 발생: {ex.Message}"
            };
            await Clients.Client(Context.ConnectionId).SendAsync("ReceiveMessage", response);
        }
    }

    public async Task Ping()
    {
        try
        {
            var message = new JObject
            {
                ["command"] = "ping",
                ["connectionId"] = Context.ConnectionId
            };
            await _messageProcessor.ProcessMessage(Context.ConnectionId, message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ping 처리 중 오류 발생: {ex.Message}");
            var response = new
            {
                status = "error",
                message = $"Ping 처리 중 오류 발생: {ex.Message}"
            };
            await Clients.Client(Context.ConnectionId).SendAsync("ReceiveMessage", response);
        }
    }
}

public class WithVueAsHost
{
    private IHost? _host;
    private readonly int _port;
    private bool _isRunning;
    private IHubContext<ChatHub>? _hubContext;
    private readonly WithFlaskAsClient _flaskClient;

    [DllImport("kernel32.dll")]
    private static extern bool AllocConsole();

    public WithVueAsHost(int port, WithFlaskAsClient flaskClient)
    {
        _port = port;
        _flaskClient = flaskClient;
        Console.WriteLine($"Vue.js 호스트 서버가 초기화되었습니다. (포트: {port})");
    }

    public IHubContext<ChatHub> GetHubContext()
    {
        if (_host == null)
        {
            throw new InvalidOperationException("서버가 초기화되지 않았습니다.");
        }
        return _host.Services.GetRequiredService<IHubContext<ChatHub>>();
    }

    public async Task SendMessageToAll(object message)
    {
        if (_hubContext == null)
        {
            _hubContext = GetHubContext();
        }

        string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine("==========================================");
        Console.WriteLine("모든 클라이언트에게 보내는 메시지:");
        Console.WriteLine(messageJson);
        Console.WriteLine("==========================================");

        await _hubContext.Clients.All.SendAsync("ReceiveMessage", message);
    }

    public async Task SendMessageToClient(string connectionId, object message)
    {
        if (_hubContext == null)
        {
            _hubContext = GetHubContext();
        }

        string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine("==========================================");
        Console.WriteLine($"클라이언트 {connectionId}에게 보내는 메시지:");
        Console.WriteLine(messageJson);
        Console.WriteLine("==========================================");

        await _hubContext.Clients.Client(connectionId).SendAsync("ReceiveMessage", message);
    }

    public async Task SendMessageToGroup(string groupName, object message)
    {
        if (_hubContext == null)
        {
            _hubContext = GetHubContext();
        }

        string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine("==========================================");
        Console.WriteLine($"그룹 {groupName}에게 보내는 메시지:");
        Console.WriteLine(messageJson);
        Console.WriteLine("==========================================");

        await _hubContext.Clients.Group(groupName).SendAsync("ReceiveMessage", message);
    }

    public async Task StartAsync()
    {
        if (_isRunning)
        {
            Console.WriteLine("서버가 이미 실행 중입니다.");
            return;
        }

        try
        {
            await _flaskClient.ConnectAsync();

            _host = Host.CreateDefaultBuilder()
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseUrls($"http://localhost:{_port}");
                    webBuilder.Configure(app =>
                    {
                        app.UseCors(builder => builder
                            .WithOrigins("http://localhost:5173") // Vue.js 개발 서버 주소
                            .AllowAnyMethod()
                            .AllowAnyHeader()
                            .AllowCredentials());

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
                            builder.WithOrigins("http://localhost:5173") // Vue.js 개발 서버 주소
                                   .AllowAnyMethod()
                                   .AllowAnyHeader()
                                   .AllowCredentials();
                        });
                    });
                    
                    services.AddSignalR(options =>
                    {
                        options.EnableDetailedErrors = true;
                        options.MaximumReceiveMessageSize = 10 * 1024 * 1024; // 10MB로 증가
                        options.KeepAliveInterval = TimeSpan.FromSeconds(30);
                        options.ClientTimeoutInterval = TimeSpan.FromSeconds(60);
                        options.HandshakeTimeout = TimeSpan.FromSeconds(30);
                    });

                    services.AddSingleton(_flaskClient);
                    services.AddSingleton<ProcessVueMessage>();
                })
                .Build();

            await _host.StartAsync();
            _isRunning = true;
            _hubContext = GetHubContext();
            
            Console.WriteLine("==========================================");
            Console.WriteLine($"Vue.js 호스트 서버가 시작되었습니다.");
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
        Console.WriteLine("서버를 종료하는 중...");
        
        if (_flaskClient != null)
        {
            await _flaskClient.DisconnectAsync();
        }
        
        if (_host != null)
        {
            await _host.StopAsync();
            _host.Dispose();
        }
        
        _isRunning = false;
        
        Console.WriteLine("==========================================");
        Console.WriteLine("서버가 종료되었습니다.");
        Console.WriteLine("==========================================");
    }
}
