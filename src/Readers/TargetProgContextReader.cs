using System;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace overlay_gpt
{
    public class TargetProgContextReader
    {
        public async Task<string> ReadFileContent(string filePath, string fileType)
        {
            string fileContent = string.Empty;
            Process? editorProcess = null;

            try
            {
                editorProcess = await StartEditorProcess(filePath, fileType);
                if (editorProcess != null)
                {
                    // 프로세스가 시작될 때까지 잠시 대기
                    editorProcess.WaitForInputIdle(5000);

                    // 파일 내용 읽기
                    var reader = ContextReaderFactory.CreateReader(AutomationElement.FromHandle(editorProcess.MainWindowHandle));
                    if (reader != null)
                    {
                        // 전체 내용 읽기
                        var (content, _, _) = reader.GetSelectedTextWithStyle(true);
                        fileContent = content;
                    }
                }
            }
            finally
            {
                await CloseEditorProcess(editorProcess);
            }

            return fileContent;
        }

        private async Task<Process?> StartEditorProcess(string filePath, string fileType)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            switch (fileType.ToLower())
            {
                case "word":
                    startInfo.FileName = "WINWORD.EXE";
                    startInfo.Arguments = $"/q /n \"{filePath}\"";
                    break;
                case "excel":
                    startInfo.FileName = "EXCEL.EXE";
                    startInfo.Arguments = $"/e \"{filePath}\"";
                    break;
                case "powerpoint":
                    startInfo.FileName = "POWERPNT.EXE";
                    startInfo.Arguments = $"/s \"{filePath}\"";
                    break;
                case "hwp":
                    startInfo.FileName = "Hwp.exe";
                    startInfo.Arguments = $"\"{filePath}\"";
                    break;
                default:
                    throw new ArgumentException($"지원하지 않는 파일 타입입니다: {fileType}");
            }

            return Process.Start(startInfo);
        }

        private async Task CloseEditorProcess(Process? process)
        {
            if (process != null && !process.HasExited)
            {
                try
                {
                    process.CloseMainWindow();
                    await Task.Delay(3000);  // 3초 대기
                    if (!process.HasExited)
                    {
                        process.Kill();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"프로세스 종료 중 오류 발생: {ex.Message}");
                }
            }
        }
    }
} 