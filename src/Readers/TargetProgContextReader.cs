using System;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Linq;

namespace overlay_gpt
{
    public class TargetProgContextReader
    {
        private BaseContextReader? _cachedReader;
        private string? _cachedFilePath;

        public async Task<(string Content, string Position, ulong? FileId, uint? VolumeId, string FileType, string FileName)> ReadFileContentAndInfo(string filePath, string fileType)
        {
            string fileContent = string.Empty;
            string position = string.Empty;
            ulong? fileId = null;
            uint? volumeId = null;
            string fileName = string.Empty;

            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                Console.WriteLine($"[{timestamp}] === TargetProgContextReader 시작 ===");
                Console.WriteLine($"[{timestamp}] 파일 경로: {filePath}, 파일 타입: {fileType}");
                
                // STA 스레드에서 실행
                var tcs = new TaskCompletionSource<(string, string, ulong?, uint?, string, string)>();
                
                var thread = new Thread(() =>
                {
                    try
                    {
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ContextReader 생성 시작...");
                        // Word 프로세스 찾기
                        var wordProcesses = Process.GetProcessesByName("WINWORD");
                        if (wordProcesses.Length == 0)
                        {
                            throw new InvalidOperationException("Word 프로세스를 찾을 수 없습니다.");
                        }
                        
                        // 가장 최근에 활성화된 Word 윈도우 찾기
                        var wordWindow = wordProcesses
                            .Select(p => AutomationElement.FromHandle(p.MainWindowHandle))
                            .FirstOrDefault(e => e != null);
                            
                        if (wordWindow == null)
                        {
                            throw new InvalidOperationException("Word 윈도우를 찾을 수 없습니다.");
                        }
                        
                        var reader = ContextReaderFactory.CreateReader(wordWindow, true, filePath);
                        
                        if (reader != null)
                        {
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ContextReader 생성 성공");
                            
                            // 파일 내용과 위치 정보 가져오기
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 내용 읽기 시작...");
                            var (content, _, pos) = reader.GetSelectedTextWithStyle(true);
                            fileContent = content;
                            position = pos;
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 내용 길이: {content?.Length ?? 0}");
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 위치 정보: {pos}");
                            
                            // 파일 정보 가져오기 (같은 reader 재사용)
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 정보 가져오기 시작...");
                            var fileInfo = reader.GetFileInfo();
                            fileId = fileInfo.FileId;
                            volumeId = fileInfo.VolumeId;
                            fileName = fileInfo.FileName;
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 정보 - FileId: {fileId}, VolumeId: {volumeId}, FileName: {fileName}");
                        }
                        else
                        {
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ContextReader 생성 실패");
                        }
                        
                        tcs.SetResult((fileContent, position, fileId, volumeId, fileType, fileName));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 내용 읽기 실패: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                        tcs.SetException(ex);
                    }
                });
                
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                
                var result = await tcs.Task;
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === TargetProgContextReader 완료 ===");
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 읽기 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                throw;
            }
        }

        // 기존 메서드들은 새로운 통합 메서드를 사용하도록 변경
        public async Task<(string Content, string Position)> ReadFileContent(string filePath, string fileType)
        {
            var result = await ReadFileContentAndInfo(filePath, fileType);
            return (result.Content, result.Position);
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo(string filePath)
        {
            // 이 메서드는 더 이상 별도의 reader를 생성하지 않음
            // 대신 ReadFileContentAndInfo를 사용하도록 호출자를 수정해야 함
            return (null, null, string.Empty, string.Empty, string.Empty);
        }
    }
} 