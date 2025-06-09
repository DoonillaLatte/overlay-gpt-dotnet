using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Microsoft.AspNetCore.SignalR;
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Network.Models.Flask;
using overlay_gpt.Services;
using overlay_gpt.Network.Models.Common;
using System.Text.Json;
using System.IO;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;
using overlay_gpt.Services;
using System.Threading;

namespace overlay_gpt.Network
{
    public class ProcessVueMessage
    {
        private readonly Dictionary<string, Func<JObject, Task>> _commandHandlers;
        private readonly IHubContext<ChatHub> _hubContext;
        private readonly WithFlaskAsClient _flaskClient;

        public ProcessVueMessage(IHubContext<ChatHub> hubContext, WithFlaskAsClient flaskClient)
        {
            _hubContext = hubContext;
            _flaskClient = flaskClient;
            _commandHandlers = new Dictionary<string, Func<JObject, Task>>
            {
                { "show_overlay", HandleShowOverlay },
                { "hide_overlay", HandleHideOverlay },
                { "update_content", HandleUpdateContent },
                { "error", HandleError },
                { "send_user_prompt", HandleSendUserPrompt },
                { "ping", HandlePing },
                { "generate_chat_id", HandleGenerateChatId },
                { "apply_response", HandleApplyResponse },
                { "apply_stored_response", HandleApplyStoredResponse },
                { "cancel_response", HandleCancelResponse },
                { "request_top_workflows", HandleRequestTopWorkflows },
                { "select_workflow", HandleSelectWorkflow }
            };
        }

        public async Task ProcessMessage(string connectionId, JObject message)
        {
            try
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 메시지 수신 시작");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ConnectionId: {connectionId}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 메시지 내용: {message.ToString(Newtonsoft.Json.Formatting.Indented)}");
                
                var command = message["command"]?.ToString();
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 명령어: {command}");

                if (string.IsNullOrEmpty(command))
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 오류: 명령어가 지정되지 않았습니다.");
                    await SendErrorResponse(connectionId, "명령어가 지정되지 않았습니다.");
                    return;
                }

                if (_commandHandlers.TryGetValue(command, out var handler))
                {
                    await handler(message);
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 명령어: {command}");
                    await SendErrorResponse(connectionId, $"처리되지 않은 명령어: {command}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
                await SendErrorResponse(connectionId, $"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleShowOverlay(JObject data)
        {
            Console.WriteLine("오버레이를 표시합니다.");
            // TODO: 실제 오버레이 표시 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleHideOverlay(JObject data)
        {
            Console.WriteLine("오버레이를 숨깁니다.");
            // TODO: 실제 오버레이 숨기기 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleUpdateContent(JObject data)
        {
            var content = data["content"]?.ToString();
            Console.WriteLine($"콘텐츠 업데이트: {content}");
            // TODO: 실제 콘텐츠 업데이트 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleError(JObject data)
        {
            var errorMessage = data["message"]?.ToString();
            Console.WriteLine($"에러 발생: {errorMessage}");
            // TODO: 실제 에러 처리 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleSendUserPrompt(JObject data)
        {
            try
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ========== HandleSendUserPrompt 시작 ==========");
                
                var vueRequest = data.ToObject<VueRequest>();
                if (vueRequest == null)
                {
                    throw new Exception("잘못된 요청 형식입니다.");
                }

                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] VueRequest 파싱 완료 - ChatId: {vueRequest.ChatId}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Prompt: {vueRequest.Prompt?.Substring(0, Math.Min(100, vueRequest.Prompt?.Length ?? 0))}...");

                var chatData = ChatDataManager.Instance.GetChatDataById(vueRequest.ChatId);
                if (chatData == null)
                {
                    chatData = new ChatData
                    {
                        ChatId = vueRequest.ChatId,
                        GeneratedTimestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                        CurrentProgram = vueRequest.CurrentProgram,
                        TargetProgram = vueRequest.TargetProgram
                    };
                    ChatDataManager.Instance.AddChatData(chatData);
                    Console.WriteLine($"해당 ChatData가 없어 새로 생성합니다. ID : {vueRequest.ChatId}");
                }

                var flaskRequest = new RequestPrompt
                {
                    ChatId = vueRequest.ChatId,
                    Prompt = vueRequest.Prompt,
                    GeneratedTimestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    RequestType = 5,
                    CurrentProgram = chatData.CurrentProgram,
                    TargetProgram = chatData.TargetProgram
                };

                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ========== Flask 요청 전송 시작 ==========");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Flask Request ChatId: {flaskRequest.ChatId}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Flask Request Command: request_prompt");
                
                await _flaskClient.EmitAsync("message", flaskRequest);
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ========== Flask 요청 전송 완료 ==========");
                Console.WriteLine("Flask 서버로 메시지를 전송했습니다.");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ========== HandleSendUserPrompt 완료 ==========");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 변환 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task HandlePing(JObject data)
        {
            try
            {
                var response = new
                {
                    status = "success",
                    message = "pong",
                    timestamp = DateTime.Now
                };
                await _hubContext.Clients.Client(data["connectionId"]?.ToString()).SendAsync("ReceiveMessage", response);
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ping 요청에 Pong 응답을 보냈습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ping 처리 중 오류 발생: {ex.Message}");
                await SendErrorResponse(data["connectionId"]?.ToString(), $"Ping 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleGenerateChatId(JObject data)
        {
            try
            {
                var chatId = data["chat_id"]?.Value<int>();
                var generatedTimestamp = data["generated_timestamp"]?.ToString(Newtonsoft.Json.Formatting.None).Trim('"');
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] JSON에서 가져온 타임스탬프: \"{generatedTimestamp}\"");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 타임스탬프 타입: {generatedTimestamp?.GetType().FullName}");

                if (chatId == null || string.IsNullOrEmpty(generatedTimestamp))
                {
                    throw new Exception("chat_id 또는 generated_timestamp가 누락되었습니다.");
                }

                // 타임스탬프 형식 정규화 (마지막 0이 사라진 경우 처리)
                else if (generatedTimestamp.EndsWith("Z") && generatedTimestamp.Split('.')[1].Length < 4)
                {
                    var parts = generatedTimestamp.Split('.');
                    var milliseconds = parts[1].TrimEnd('Z').PadRight(3, '0');
                    generatedTimestamp = $"{parts[0]}.{milliseconds}Z";
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 타임스탬프 형식 보정: {generatedTimestamp}");
                }

                var chatData = ChatDataManager.Instance.GetChatDataByTimeStamp(generatedTimestamp);
                if (chatData != null)
                {
                    chatData.ChatId = chatId.Value;
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ChatData ID 업데이트 완료: {chatId} (Timestamp: {generatedTimestamp})");
                }
                else
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 해당 Timestamp({generatedTimestamp})에 일치하는 ChatData를 찾을 수 없습니다.");
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chat ID 생성 처리 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task HandleApplyResponse(JObject data)
        {
            try
            {
                // Vue에서 전송한 apply_content 확인
                var applyContent = data["apply_content"]?.ToString();
                var chatId = data["chat_id"]?.Value<int>() ?? 0;
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Vue에서 전송한 apply_content 길이: {applyContent?.Length ?? 0}");
                
                // 해당 ChatID를 통해 데이터 가져오기
                var chatData = ChatDataManager.Instance.GetChatDataById(chatId);
                if (chatData == null)
                {
                    throw new Exception("해당 ChatID를 통해 데이터를 가져오지 못했습니다.");
                }
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 채팅 데이터 정보:");
                Console.WriteLine($"- ChatID: {chatData.ChatId}");
                Console.WriteLine($"- 현재 프로그램: {chatData.CurrentProgram?.FileType} - {chatData.CurrentProgram?.FileName}");
                Console.WriteLine($"- 대상 프로그램: {chatData.TargetProgram?.FileType} - {chatData.TargetProgram?.FileName}");
                
                ProgramInfo programToChange = null;
                bool isTargetProg = false;
                
                // target_program이 null인지 확인
                if (chatData.TargetProgram == null)
                {
                    // null이라면 current_program에 생성된 context를 적용
                    programToChange = chatData.CurrentProgram;
                    Console.WriteLine("대상 프로그램이 null이므로 현재 프로그램에 적용합니다.");
                    isTargetProg = false;
                }
                else 
                {
                    // null이 아니라면 target_program에 생성된 context를 적용
                    programToChange = chatData.TargetProgram;
                    Console.WriteLine("대상 프로그램에 적용합니다.");
                    isTargetProg = true;
                }

                // 프로그램 정보가 null인지 확인
                if (programToChange == null)
                {
                    throw new Exception("적용할 프로그램 정보가 없습니다. 프로그램을 선택해주세요.");
                }

                // 파일 타입이 유효한지 확인
                if (string.IsNullOrEmpty(programToChange.FileType))
                {
                    throw new Exception("프로그램 타입이 지정되지 않았습니다.");
                }

                // Vue에서 전송한 원본 HTML이 있으면 사용, 없으면 기존 GeneratedContext 사용
                string contextToApply = programToChange.GeneratedContext;
                
                Console.WriteLine($"적용할 컨텍스트 길이: {contextToApply?.Length ?? 0} 문자");
                Console.WriteLine($"적용할 위치: {programToChange.Position}");
                Console.WriteLine($"사용된 컨텍스트 소스: {(applyContent != null ? "Vue 원본 HTML" : "기존 GeneratedContext")}");
                
                // 해당 프로그램에 적용 (변수를 미리 캡처)
                var contextToApplyLocal = contextToApply;
                var programToChangeLocal = programToChange;
                var isTargetProgLocal = isTargetProg;
                
                await Task.Run(() =>
                {
                    var thread = new Thread(() =>
                    {
                        try
                        {
                            // 한글 프로세스 실행 여부 확인
                            if (programToChangeLocal.FileType == "Hwp" && Process.GetProcessesByName("Hwp").Length == 0)
                            {
                                throw new Exception("한글(HWP)이 실행되어 있지 않습니다. 한글을 실행한 후 다시 시도해주세요.");
                            }

                            var writer = ContextWriterFactory.CreateWriter(programToChangeLocal.FileType);
                            
                            if (writer == null)
                            {
                                throw new Exception("지원하지 않는 프로그램입니다.");
                            }
                            Console.WriteLine($"Writer 생성 완료: {programToChangeLocal.FileType}");

                            // 생성된 컨텍스트가 있는지 확인
                            if (contextToApplyLocal == null)
                            {
                                throw new Exception("적용할 컨텍스트가 없습니다.");
                            }
                            
                            writer.IsTargetProg = isTargetProgLocal;

                            // 파일 열기 시도
                            Console.WriteLine($"파일 열기 시도: {programToChangeLocal.FilePath}");
                            if (!writer.OpenFile(programToChangeLocal.FilePath))
                            {
                                throw new Exception("파일을 열 수 없습니다. 파일이 존재하는지 확인해주세요.");
                            }
                            Console.WriteLine("파일 열기 성공");

                            // 컨텍스트 적용
                            Console.WriteLine($"컨텍스트 적용 시작... (길이: {contextToApplyLocal.Length})");
                            bool success = writer.ApplyTextWithStyle(
                                contextToApplyLocal,
                                programToChangeLocal.Position
                            );

                            if (!success)
                            {
                                throw new Exception("컨텍스트 적용에 실패했습니다.");
                            }

                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 컨텍스트 적용 완료");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 컨텍스트 적용 중 오류 발생: {ex.Message}");
                            Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                            if (ex.InnerException != null)
                            {
                                Console.WriteLine($"내부 예외: {ex.InnerException.Message}");
                                Console.WriteLine($"내부 예외 스택 트레이스: {ex.InnerException.StackTrace}");
                            }
                            throw;
                        }
                    });

                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                });

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 컨텍스트 적용 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"내부 예외: {ex.InnerException.Message}");
                    Console.WriteLine($"내부 예외 스택 트레이스: {ex.InnerException.StackTrace}");
                }
                throw;
            }
        }

        private async Task HandleApplyStoredResponse(JObject data)
        {
            try
            {
                var chatId = data["chat_id"]?.Value<int>() ?? -1;
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] HandleApplyStoredResponse 시작 - ChatID: {chatId}");

                // 해당 ChatID를 통해 데이터 가져오기
                var chatData = ChatDataManager.Instance.GetChatDataById(chatId);
                if (chatData == null)
                {
                    throw new Exception("해당 ChatID를 통해 데이터를 가져오지 못했습니다.");
                }
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 채팅 데이터 정보:");
                Console.WriteLine($"- ChatID: {chatData.ChatId}");
                Console.WriteLine($"- 현재 프로그램: {chatData.CurrentProgram?.FileType} - {chatData.CurrentProgram?.FileName}");
                Console.WriteLine($"- 대상 프로그램: {chatData.TargetProgram?.FileType} - {chatData.TargetProgram?.FileName}");
                Console.WriteLine($"- 저장된 적용용 컨텍스트 길이: {chatData.DotnetApplyContext?.Length ?? 0}");
                
                // 적용할 프로그램과 컨텍스트 결정
                ProgramInfo programToChange = null;
                bool isTargetProg = chatData.TargetProgram != null;
                
                if (isTargetProg)
                {
                    programToChange = chatData.TargetProgram;
                    Console.WriteLine("대상 프로그램에 적용합니다.");
                }
                else
                {
                    programToChange = chatData.CurrentProgram;
                    Console.WriteLine("현재 프로그램에 적용합니다.");
                }

                if (programToChange == null)
                {
                    throw new Exception("적용할 프로그램을 찾을 수 없습니다.");
                }

                // 저장된 dotnet 적용용 컨텍스트 사용
                var contextToApply = chatData.DotnetApplyContext;
                if (string.IsNullOrEmpty(contextToApply))
                {
                    throw new Exception("적용할 컨텍스트가 없습니다.");
                }

                Console.WriteLine($"사용할 컨텍스트 길이: {contextToApply.Length}");
                
                // 변수를 미리 캡처 (Thread 클로저 문제 해결)
                var contextToApplyLocal = contextToApply;
                var programToChangeLocal = programToChange;
                var isTargetProgLocal = isTargetProg;
                
                // 해당 프로그램에 적용
                await Task.Run(() =>
                {
                    var thread = new Thread(() =>
                    {
                        try
                        {
                            // 한글 프로세스 실행 여부 확인
                            if (programToChangeLocal.FileType == "Hwp" && Process.GetProcessesByName("Hwp").Length == 0)
                            {
                                throw new Exception("한글(HWP)이 실행되어 있지 않습니다. 한글을 실행한 후 다시 시도해주세요.");
                            }

                            var writer = ContextWriterFactory.CreateWriter(programToChangeLocal.FileType);
                            
                            if (writer == null)
                            {
                                throw new Exception("지원하지 않는 프로그램입니다.");
                            }
                            Console.WriteLine($"Writer 생성 완료: {programToChangeLocal.FileType}");

                            writer.IsTargetProg = isTargetProgLocal;

                            // 파일 열기 시도
                            Console.WriteLine($"파일 열기 시도: {programToChangeLocal.FilePath}");
                            if (!writer.OpenFile(programToChangeLocal.FilePath))
                            {
                                throw new Exception("파일을 열 수 없습니다. 파일이 존재하는지 확인해주세요.");
                            }
                            Console.WriteLine("파일 열기 성공");

                            // 컨텍스트 적용
                            Console.WriteLine($"저장된 컨텍스트 적용 시작... (길이: {contextToApplyLocal.Length})");
                            bool success = writer.ApplyTextWithStyle(
                                contextToApplyLocal,
                                programToChangeLocal.Position
                            );

                            if (success)
                            {
                                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 저장된 컨텍스트 적용 성공!");
                            }
                            else
                            {
                                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 저장된 컨텍스트 적용 실패");
                                throw new Exception("컨텍스트 적용에 실패했습니다.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 적용 스레드 오류: {ex.Message}");
                            throw;
                        }
                    });
                    
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                });
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] HandleApplyStoredResponse 완료");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] HandleApplyStoredResponse 오류: {ex.Message}");
                throw;
            }
        }

        private async Task HandleCancelResponse(JObject data)
        {
            // TODO: 응답 취소 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleRequestTopWorkflows(JObject data)
        {
            try
            {
                var chatId = data["chat_id"]?.Value<int>();
                var fileType = data["file_type"]?.ToString();
                if (chatId == null)
                {
                    throw new Exception("chat_id가 누락되었습니다.");
                }

                var chatData = ChatDataManager.Instance.GetChatDataById(chatId.Value);
                if (chatData == null)
                {
                    throw new Exception($"Chat ID {chatId}에 해당하는 데이터를 찾을 수 없습니다.");
                }

                var flaskRequest = new
                {
                    command = "get_workflows",
                    chat_id = chatId,
                    file_type = fileType,
                    current_program = chatData.CurrentProgram
                };

                await _flaskClient.EmitAsync("message", flaskRequest);
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Flask 서버로 워크플로우 요청을 전송했습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"워크플로우 요청 처리 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task HandleSelectWorkflow(JObject data)
        {
            try
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 워크플로우 선택 프로세스 시작");
                
                var chatId = data["chat_id"]?.Value<int>();
                var fileType = data["file_type"]?.ToString();
                var targetProgram = data["target_program"]?.ToObject<string[]>();

                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 파라미터 확인 - ChatId: {chatId}, FileType: {fileType}, TargetProgram 길이: {targetProgram?.Length}");

                if (chatId == null || string.IsNullOrEmpty(fileType) || targetProgram == null || targetProgram.Length < 2)
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 오류: 필수 파라미터 누락");
                    throw new Exception("필수 파라미터가 누락되었습니다.");
                }

                var chatData = ChatDataManager.Instance.GetChatDataById(chatId.Value);
                if (chatData == null)
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 오류: Chat ID {chatId}에 해당하는 데이터를 찾을 수 없음");
                    throw new Exception($"Chat ID {chatId}에 해당하는 데이터를 찾을 수 없습니다.");
                }

                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ChatData 조회 성공");

                var filePath = targetProgram[1];
                
                // 테스트 파일 경로
                //filePath = "C:\\Users\\beste\\OneDrive\\Desktop\\testFolder\\single_test.pptx";
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 대상 파일 경로: {filePath}");

                // 파일 내용과 정보를 한 번에 가져오기 (중복 호출 방지)
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 파일 내용 및 정보 읽기 시작");
                var contextReader = new TargetProgContextReader();
                var (fileContent, position, fileId, volumeId, readFileType, fileName) = await contextReader.ReadFileContentAndInfo(filePath, fileType);
                
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 파일 내용 읽기 완료 - 크기: {fileContent.Length} 문자");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 위치 정보: {position}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 파일 정보 - FileId: {fileId}, VolumeId: {volumeId}, FileName: {fileName}");

                if (fileId == null || volumeId == null)
                {
                    throw new Exception("파일 정보를 가져올 수 없습니다.");
                }

                // ChatData의 target_program 업데이트
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] TargetProgram 설정 시작");
                Console.WriteLine($"- FilePath: {filePath}");
                Console.WriteLine($"- FileType: {fileType}");
                Console.WriteLine($"- FileId: {fileId.Value}");
                Console.WriteLine($"- VolumeId: {volumeId.Value}");
                Console.WriteLine($"- GeneratedContext : {fileContent}");
                Console.WriteLine($"- Position: {position}");
                chatData.TargetProgram = new ProgramInfo
                {
                    FileName = fileName,
                    FilePath = filePath,
                    FileType = fileType,
                    FileId = fileId.Value,
                    VolumeId = volumeId.Value,
                    Context = fileContent,
                    Position = position
                };
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] TargetProgram 설정 완료");

                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 워크플로우 선택 프로세스 완료");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 워크플로우 선택 처리 중 오류 발생: {ex.Message}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 스택 트레이스: {ex.StackTrace}");
                throw;
            }
        }

        private async Task SendErrorResponse(string connectionId, string errorMessage)
        {
            var response = new
            {
                status = "error",
                message = errorMessage
            };
            await _hubContext.Clients.Client(connectionId).SendAsync("ReceiveMessage", response);
        }

        // 새로운 명령어 핸들러를 추가하는 메서드
        public void RegisterCommandHandler(string command, Func<JObject, Task> handler)
        {
            if (_commandHandlers.ContainsKey(command))
            {
                _commandHandlers[command] = handler;
            }
            else
            {
                _commandHandlers.Add(command, handler);
            }
        }
    }
}
