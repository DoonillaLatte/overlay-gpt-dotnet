using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SocketIOClient;
using Newtonsoft.Json.Linq;
using overlay_gpt.Network.Models.Common;
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Services;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;

namespace overlay_gpt.Network
{
    public class ProcessFlaskMessage
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int NtCreateFile(
            out IntPtr FileHandle,
            uint DesiredAccess,
            ref OBJECT_ATTRIBUTES ObjectAttributes,
            out IO_STATUS_BLOCK IoStatusBlock,
            IntPtr AllocationSize,
            uint FileAttributes,
            uint ShareAccess,
            uint CreateDisposition,
            uint CreateOptions,
            IntPtr EaBuffer,
            uint EaLength);

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int NtQueryInformationFile(
            IntPtr FileHandle,
            out IO_STATUS_BLOCK IoStatusBlock,
            IntPtr FileInformation,
            uint Length,
            int FileInformationClass);

        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint dwVolumeSerialNumber;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint nNumberOfLinks;
            public uint nFileIndexHigh;
            public uint nFileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct OBJECT_ATTRIBUTES
        {
            public int Length;
            public IntPtr RootDirectory;
            public IntPtr ObjectName;
            public uint Attributes;
            public IntPtr SecurityDescriptor;
            public IntPtr SecurityQualityOfService;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IO_STATUS_BLOCK
        {
            public uint Status;
            public IntPtr Information;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILE_ID_BOTH_DIR_INFO
        {
            public uint NextEntryOffset;
            public uint FileIndex;
            public long CreationTime;
            public long LastAccessTime;
            public long LastWriteTime;
            public long ChangeTime;
            public long EndOfFile;
            public long AllocationSize;
            public uint FileAttributes;
            public uint FileNameLength;
            public uint EaSize;
            public byte ShortNameLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 12)]
            public byte[] ShortName;
            public long FileId;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] FileName;
        }

        private const uint GENERIC_READ = 0x80000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        private const uint FILE_OPEN = 1;
        private const uint FILE_OPEN_BY_FILE_ID = 0x00002000;
        private const uint FILE_READ_ATTRIBUTES = 0x0080;
        private const uint FILE_SHARE_DELETE = 0x00000004;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
        private const uint OBJ_CASE_INSENSITIVE = 0x00000040;
        private const int FileIdBothDirectoryInformation = 37;

        private readonly Dictionary<string, Func<JObject, Task>> _commandHandlers;
        private readonly NtfsFileFinder _fileFinder;

        public ProcessFlaskMessage()
        {
            _commandHandlers = new Dictionary<string, Func<JObject, Task>>
            {
                { "show_overlay", HandleShowOverlay },
                { "hide_overlay", HandleHideOverlay },
                { "update_content", HandleUpdateContent },
                { "error", HandleError },
                { "generated_response", HandleGeneratedResponse },
                { "response_workflows", HandleResponseWorkflows }
            };
            _fileFinder = new NtfsFileFinder();
        }

        public async Task ProcessMessage(SocketIOResponse response)
        {
            try
            {
                Console.WriteLine("ProcessMessage 시작");
                var jsonString = response.GetValue<object>().ToString();
                var jsonData = JObject.Parse(jsonString);
                Console.WriteLine($"수신된 메시지: {jsonData}");
                
                var command = jsonData["command"]?.ToString();
                Console.WriteLine($"명령어: {command}");

                if (string.IsNullOrEmpty(command))
                {
                    Console.WriteLine("명령어가 지정되지 않았습니다.");
                    return;
                }

                Console.WriteLine($"사용 가능한 핸들러 목록: {string.Join(", ", _commandHandlers.Keys)}");
                
                if (_commandHandlers.TryGetValue(command, out var handler))
                {
                    Console.WriteLine($"핸들러 실행: {command}");
                    await handler(jsonData);
                    Console.WriteLine($"핸들러 실행 완료: {command}");
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 명령어: {command}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            }
        }

        private async Task HandleShowOverlay(JObject data)
        {
            // 오버레이 표시 로직 구현
            Console.WriteLine("오버레이를 표시합니다.");
            // TODO: 실제 오버레이 표시 로직 구현
        }

        private async Task HandleHideOverlay(JObject data)
        {
            // 오버레이 숨기기 로직 구현
            Console.WriteLine("오버레이를 숨깁니다.");
            // TODO: 실제 오버레이 숨기기 로직 구현
        }

        private async Task HandleUpdateContent(JObject data)
        {
            var content = data["content"]?.ToString();
            Console.WriteLine($"콘텐츠 업데이트: {content}");
            // TODO: 실제 콘텐츠 업데이트 로직 구현
        }

        private async Task HandleError(JObject data)
        {
            var errorMessage = data["message"]?.ToString();
            Console.WriteLine($"에러 발생: {errorMessage}");
            // TODO: 실제 에러 처리 로직 구현
        }

        private async Task HandleGeneratedResponse(JObject data)
        {
            try
            {
                Console.WriteLine("HandleGeneratedResponse 시작");
                var chatId = data["chat_id"]?.Value<int>() ?? -1;
                var title = data["title"]?.ToString();
                var message = data["message"]?.ToString();
                var status = data["status"]?.ToString();

                Console.WriteLine($"받은 데이터 - chatId: {chatId}, message: {message}, status: {status}");

                if (string.IsNullOrEmpty(message))
                {
                    Console.WriteLine("message가 비어있습니다.");
                    return;
                }

                // HTML 형식인지 확인
                bool isHtml = message.Contains("<") && message.Contains(">");
                string textType = isHtml ? "text_to_apply" : "text_plain";

                var chatData = Services.ChatDataManager.Instance.GetChatDataById(chatId);
                if (chatData == null)
                {
                    Console.WriteLine($"chat_id {chatId}에 해당하는 ChatData를 찾을 수 없습니다.");
                    return;
                }

                if(chatData.Title != title) 
                {
                    chatData.Title = title;
                }
                chatData.Texts.Add(new TextData { Type = textType, Content = message });
                if(chatData.TargetProgram == null)
                {
                    if (chatData.CurrentProgram != null) // CurrentProgram에 대한 null 체크 추가
                    {
                        chatData.CurrentProgram.Context = message;
                        chatData.CurrentProgram.GeneratedContext = message;
                    }
                    else
                    {
                        Console.WriteLine($"chatId {chatId}의 CurrentProgram이 null입니다.");
                    }
                }
                else
                {
                    if (chatData.TargetProgram != null) // TargetProgram에 대한 null 체크 추가
                    {
                        chatData.TargetProgram.Context = message;
                        chatData.TargetProgram.GeneratedContext = message;
                    }
                    else
                    {
                        Console.WriteLine($"chatId {chatId}의 TargetProgram이 null입니다.");
                    }
                }
                Console.WriteLine($"ChatData {chatId}에 메시지가 추가되었습니다.");

                // Vue로 display_text 메시지 전송
                var displayText = new DisplayText
                {
                    ChatId = chatId,
                    Title = title,
                    GeneratedTimestamp = chatData.GeneratedTimestamp,
                    CurrentProgram = chatData.CurrentProgram,
                    TargetProgram = chatData.TargetProgram,
                    Texts = new List<TextData> { new TextData { Type = textType, Content = message } }
                };

                var vueServer = MainWindow.Instance.VueServer;
                if (vueServer != null)
                {
                    await vueServer.SendMessageToAll(displayText);
                    Console.WriteLine($"Vue로 display_text 메시지 전송 완료: chat_id {chatId}");
                }
                else
                {
                    Console.WriteLine("Vue 서버가 초기화되지 않았습니다.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleResponseWorkflows(JObject data)
        {
            try
            {
                Console.WriteLine("HandleResponseWorkflows 시작");
                var chatId = data["chat_id"]?.Value<int>() ?? -1;
                var similarProgramIds = new List<List<long>>();
                try 
                {
                    var rawIds = data["similar_program_ids"]?.ToObject<List<List<object>>>();
                    if (rawIds != null)
                    {
                        similarProgramIds = rawIds.Select(innerList => 
                            innerList.Select(item => 
                                item != null ? Convert.ToInt64(item) : 0L
                            ).ToList()
                        ).ToList();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"similar_program_ids 변환 중 오류: {ex.Message}");
                }
                var status = data["status"]?.ToString();
                string fileType = data["file_type"]?.ToString();

                Console.WriteLine($"받은 데이터 - chatId: {chatId}, similarProgramIds: {string.Join(", ", similarProgramIds.Select(x => $"[{x[0]}, {x[1]}]"))}");

                var chatData = Services.ChatDataManager.Instance.GetChatDataById(chatId);
                if (chatData == null)
                {
                    Console.WriteLine($"chat_id {chatId}에 해당하는 ChatData를 찾을 수 없습니다.");
                    return;
                }

                // 파일 정보를 찾아서 변환
                var convertedPrograms = new List<List<string>>();

                foreach (var programId in similarProgramIds)
                {
                    var fileId = programId[0];
                    var volumeId = programId[1];

                    var foundFile = _fileFinder.FindFileByFileIdAndVolumeId(fileId, volumeId);
                    if (foundFile != null)
                    {
                        var fileName = Path.GetFileName(foundFile);
                        var filePath = foundFile;
                        
                        
                        

                        convertedPrograms.Add(new List<string> { fileName, filePath });
                    }
                }
                
                // 임시 파일 경로 생성. 나중에 지울 것
                convertedPrograms.Add(new List<string> { "임시파일1.확장자", "임시파일경로1/임시파일경로2/임시파일1.확장자" });
                convertedPrograms.Add(new List<string> { "임시파일2.확장자", "임시파일경로1/임시파일경로2/임시파일2.확장자" });
                convertedPrograms.Add(new List<string> { "임시파일3.확장자", "임시파일경로1/임시파일경로2/임시파일3.확장자" });

                // Vue로 메시지 전송
                var responseData = new
                {
                    command = "response_top_workflows",
                    chat_id = chatId,
                    file_type = fileType,
                    similar_programs = convertedPrograms,
                    status = status
                };

                var vueServer = MainWindow.Instance.VueServer;
                if (vueServer != null)
                {
                    await vueServer.SendMessageToAll(responseData);
                    Console.WriteLine($"Vue로 response_top_workflows 메시지 전송 완료: chat_id {chatId}");
                }
                else
                {
                    Console.WriteLine("Vue 서버가 초기화되지 않았습니다.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"워크플로우 응답 처리 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            }
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
