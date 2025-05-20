using Microsoft.AspNetCore.SignalR;
using System.Text.Json;
using System.Threading.Tasks;
using System.Threading;

namespace overlay_gpt.Network;

public class ChatHub : Hub
{
    private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
    private static readonly Dictionary<string, DateTime> _lastPingTime = new();

    public override async Task OnConnectedAsync()
    {
        await _semaphore.WaitAsync();
        try
        {
            _lastPingTime[Context.ConnectionId] = DateTime.Now;
            Console.WriteLine("==========================================");
            Console.WriteLine($"새로운 클라이언트가 연결되었습니다.");
            Console.WriteLine($"연결 ID: {Context.ConnectionId}");
            Console.WriteLine($"연결 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine("==========================================");
            await base.OnConnectedAsync();
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public override async Task OnDisconnectedAsync(Exception? exception)
    {
        await _semaphore.WaitAsync();
        try
        {
            _lastPingTime.Remove(Context.ConnectionId);
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
        finally
        {
            _semaphore.Release();
        }
    }

    public async Task Ping()
    {
        await _semaphore.WaitAsync();
        try
        {
            _lastPingTime[Context.ConnectionId] = DateTime.Now;
            await Clients.Caller.SendAsync("ReceiveMessage", new { status = "success", message = "pong" });
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public async Task SendMessage(object message)
    {
        await _semaphore.WaitAsync();
        try
        {
            // 메시지 로깅
            string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine($"받은 메시지:\n{messageJson}");

            // 메시지 처리
            var response = new
            {
                status = "success",
                message = "메시지를 성공적으로 처리했습니다."
            };

            await Clients.Caller.SendAsync("ReceiveMessage", response);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            
            var errorResponse = new
            {
                status = "error",
                message = $"메시지 처리 중 오류가 발생했습니다: {ex.Message}"
            };

            await Clients.Caller.SendAsync("ReceiveMessage", errorResponse);
        }
        finally
        {
            _semaphore.Release();
        }
    }
} 