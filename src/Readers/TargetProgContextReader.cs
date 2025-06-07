using System;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;

namespace overlay_gpt
{
    public class TargetProgContextReader
    {
        public async Task<(string Content, string Position)> ReadFileContent(string filePath, string fileType)
        {
            string fileContent = string.Empty;
            string position = string.Empty;

            try
            {
                Console.WriteLine($"[DEBUG] 파일 경로: {filePath}, 파일 타입: {fileType}");
                
                // STA 스레드에서 실행
                var tcs = new TaskCompletionSource<(string, string)>();
                
                var thread = new Thread(() =>
                {
                    try
                    {
                        var reader = ContextReaderFactory.CreateReader(AutomationElement.FromHandle(Process.GetCurrentProcess().MainWindowHandle), true, filePath);
                        
                        if (reader != null)
                        {
                            Console.WriteLine("[DEBUG] ContextReader 생성 성공");
                            var (content, _, pos) = reader.GetSelectedTextWithStyle(true);
                            fileContent = content;
                            position = pos;
                            Console.WriteLine($"[DEBUG] 파일 내용 길이: {content?.Length ?? 0}");
                            Console.WriteLine($"[DEBUG] 위치 정보: {pos}");
                        }
                        else
                        {
                            Console.WriteLine("[DEBUG] ContextReader 생성 실패");
                        }
                        
                        tcs.SetResult((fileContent, position));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ERROR] 파일 내용 읽기 실패: {ex.Message}");
                        tcs.SetException(ex);
                    }
                });
                
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                
                var result = await tcs.Task;
                fileContent = result.Item1;
                position = result.Item2;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] 파일 읽기 중 오류 발생: {ex.Message}");
                Console.WriteLine($"[ERROR] 스택 트레이스: {ex.StackTrace}");
                throw;
            }

            return (fileContent, position);
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo(string filePath)
        {
            try
            {
                var reader = ContextReaderFactory.CreateReader(AutomationElement.FromHandle(Process.GetCurrentProcess().MainWindowHandle), true, filePath);
                if (reader != null)
                {
                    return reader.GetFileInfo();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] 파일 정보 가져오기 실패: {ex.Message}");
            }
            return (null, null, string.Empty, string.Empty, string.Empty);
        }
    }
} 