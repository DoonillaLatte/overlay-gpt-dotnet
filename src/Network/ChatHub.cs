using Microsoft.AspNetCore.SignalR;
using System.Text.Json;
using System.Threading.Tasks;
using System.Threading;
using overlay_gpt.Network.Models;

namespace overlay_gpt.Network;

public class ChatHub : Hub
{
    private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
    private static readonly Dictionary<string, DateTime> _lastPingTime = new();
    private readonly SocketIOConnection _socketIOConnection;

    public ChatHub(SocketIOConnection socketIOConnection)
    {
        _socketIOConnection = socketIOConnection;
    }

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
        Console.WriteLine("==========================================");
        Console.WriteLine("SendMessage 호출됨");
        Console.WriteLine("==========================================");
        try
        {
            // 메시지 로깅
            string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine("==========================================");
            Console.WriteLine("Vue에서 받은 메시지:");
            Console.WriteLine(messageJson);
            Console.WriteLine("==========================================");

            // JSON 메시지 파싱
            var jsonElement = JsonSerializer.SerializeToElement(message);
            if (!jsonElement.TryGetProperty("command", out var commandElement))
            {
                throw new InvalidOperationException("메시지에 'command' 필드가 없습니다.");
            }

            string command = commandElement.GetString() ?? throw new InvalidOperationException("command 값이 null입니다.");
            object? response;

            // command에 따른 처리
            switch (command.ToLower())
            {
                case "ping":
                    response = new { status = "success", message = "pong" };
                    break;

                case "get_status":
                    response = new { status = "success", data = new { isRunning = true } };
                    break;

                case "send_user_prompt":
                    if (!jsonElement.TryGetProperty("chat_id", out var chatIdElement) ||
                        !jsonElement.TryGetProperty("prompt", out var promptElement))
                    {
                        throw new InvalidOperationException("필수 필드가 누락되었습니다: chat_id 또는 prompt");
                    }

                    int chatId = chatIdElement.GetInt32();
                    string prompt = promptElement.GetString() ?? throw new InvalidOperationException("prompt 값이 null입니다.");

                    // Flask 요청 형식으로 변환
                    var flaskRequest = new FlaskRequest
                    {
                        ChatId = chatId,
                        Prompt = prompt,
                        RequestType = jsonElement.TryGetProperty("request_type", out var requestTypeElement) 
                            ? requestTypeElement.GetInt32() 
                            : 1,
                        Description = jsonElement.TryGetProperty("description", out var descriptionElement) 
                            ? descriptionElement.GetString() ?? string.Empty 
                            : string.Empty,
                        CurrentProgram = jsonElement.TryGetProperty("current_program", out var currentProgramElement) 
                            ? JsonSerializer.Deserialize<ProgramInfo>(currentProgramElement.GetRawText()) ?? new ProgramInfo()
                            : new ProgramInfo(),
                        TargetProgram = jsonElement.TryGetProperty("target_program", out var targetProgramElement) 
                            ? JsonSerializer.Deserialize<ProgramInfo>(targetProgramElement.GetRawText()) ?? new ProgramInfo()
                            : new ProgramInfo()
                    };

                    // Flask 서버로 전송
                    string flaskRequestJson = JsonSerializer.Serialize(flaskRequest, new JsonSerializerOptions { WriteIndented = true });
                    Console.WriteLine("==========================================");
                    Console.WriteLine("Flask로 보내는 메시지:");
                    Console.WriteLine(flaskRequestJson);
                    Console.WriteLine("==========================================");
                    await _socketIOConnection.SendMessageAsync(flaskRequestJson);

                    response = new { status = "success", message = "프롬프트가 Flask 서버로 전송되었습니다." };
                    break;

                case "request_single_generated_response":
                    var singleResponseRequest = new FlaskRequest
                    {
                        ChatId = jsonElement.TryGetProperty("chat_id", out var singleChatIdElement) 
                            ? singleChatIdElement.GetInt32() 
                            : 0,
                        Prompt = jsonElement.TryGetProperty("prompt", out var singlePromptElement) 
                            ? singlePromptElement.GetString() ?? string.Empty 
                            : string.Empty,
                        RequestType = jsonElement.TryGetProperty("request_type", out var singleReqTypeElement) 
                            ? singleReqTypeElement.GetInt32() 
                            : 1,
                        Description = jsonElement.TryGetProperty("description", out var singleDescElement) 
                            ? singleDescElement.GetString() ?? string.Empty 
                            : string.Empty,
                        CurrentProgram = jsonElement.TryGetProperty("current_program", out var singleCurrProgElement) 
                            ? JsonSerializer.Deserialize<ProgramInfo>(singleCurrProgElement.GetRawText()) ?? new ProgramInfo()
                            : new ProgramInfo(),
                        TargetProgram = jsonElement.TryGetProperty("target_program", out var singleTargetProgElement) 
                            ? JsonSerializer.Deserialize<ProgramInfo>(singleTargetProgElement.GetRawText()) ?? new ProgramInfo()
                            : new ProgramInfo()
                    };

                    // Flask 서버로 전송
                    string singleRequestJson = JsonSerializer.Serialize(singleResponseRequest, new JsonSerializerOptions { WriteIndented = true });
                    Console.WriteLine("==========================================");
                    Console.WriteLine("Flask로 보내는 메시지:");
                    Console.WriteLine(singleRequestJson);
                    Console.WriteLine("==========================================");
                    await _socketIOConnection.SendMessageAsync(singleRequestJson);

                    response = new { status = "success", message = "프롬프트가 Flask 서버로 전송되었습니다." };
                    break;

                default:
                    response = new { status = "error", message = $"알 수 없는 명령어: {command}" };
                    break;
            }

            string responseJson = JsonSerializer.Serialize(response, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine("==========================================");
            Console.WriteLine("Vue로 보내는 응답:");
            Console.WriteLine(responseJson);
            Console.WriteLine("==========================================");
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

            string errorJson = JsonSerializer.Serialize(errorResponse, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine("==========================================");
            Console.WriteLine("Vue로 보내는 에러 응답:");
            Console.WriteLine(errorJson);
            Console.WriteLine("==========================================");
            await Clients.Caller.SendAsync("ReceiveMessage", errorResponse);
        }
        finally
        {
            _semaphore.Release();
        }
    }
} 