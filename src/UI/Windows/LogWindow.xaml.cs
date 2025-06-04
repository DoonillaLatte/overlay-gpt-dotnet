using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Controls;
using System.Text.Json;
using System.Threading;
using overlay_gpt.Network;
using overlay_gpt.Network.Models;
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Network.Models.Common;
using overlay_gpt.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace overlay_gpt;

public partial class LogWindow : Window
{
    private static LogWindow? _instance;
    private static readonly SemaphoreSlim _instanceSemaphore = new SemaphoreSlim(1, 1);
    private static readonly SemaphoreSlim _logSemaphore = new SemaphoreSlim(1, 1);
    private WithFlaskAsClient? _flaskClient;
    private WithVueAsHost? _vueHost;
    private readonly NtfsFileFinder _ntfsFileFinder;

    public static LogWindow Instance
    {
        get
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new LogWindow();
            }
            return _instance;
        }
    }

    public static async Task<LogWindow> GetInstanceAsync()
    {
        await _instanceSemaphore.WaitAsync();
        try
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new LogWindow();
            }
            return _instance;
        }
        finally
        {
            _instanceSemaphore.Release();
        }
    }

    public LogWindow()
    {
        InitializeComponent();
        _ntfsFileFinder = new NtfsFileFinder();
    }

    public void SetServers(WithFlaskAsClient flaskClient, WithVueAsHost vueHost)
    {
        _flaskClient = flaskClient;
        _vueHost = vueHost;
        Log("서버 연결이 설정되었습니다.");
    }

    private void ApplyFontStyle(Run run, Dictionary<string, object> styleAttributes)
    {
        // 폰트 패밀리 적용
        if (styleAttributes.TryGetValue("FontName", out var fontName) && fontName != null)
        {
            try
            {
                run.FontFamily = new FontFamily(fontName.ToString());
            }
            catch
            {
                // 폰트 변환 실패 시 기본 폰트 사용
            }
        }

        // 폰트 크기 적용
        if (styleAttributes.TryGetValue("FontSize", out var fontSize) && fontSize != null)
        {
            try
            {
                run.FontSize = Convert.ToDouble(fontSize);
            }
            catch
            {
                // 폰트 크기 변환 실패 시 기본 크기 사용
            }
        }

        // 폰트 굵기 적용
        if (styleAttributes.TryGetValue("FontWeight", out var fontWeight) && fontWeight != null)
        {
            try
            {
                var weight = Convert.ToDouble(fontWeight);
                if (weight >= 700)
                    run.FontWeight = FontWeights.Bold;
                else if (weight >= 600)
                    run.FontWeight = FontWeights.SemiBold;
                else if (weight >= 500)
                    run.FontWeight = FontWeights.Medium;
                else if (weight >= 400)
                    run.FontWeight = FontWeights.Regular;
                else if (weight >= 300)
                    run.FontWeight = FontWeights.Light;
                else
                    run.FontWeight = FontWeights.Thin;
            }
            catch
            {
                // 폰트 굵기 변환 실패 시 기본 굵기 사용
            }
        }

        // 이탤릭 스타일 적용
        if (styleAttributes.TryGetValue("FontStyle", out var fontStyle) && fontStyle != null)
        {
            try
            {
                var style = Convert.ToInt32(fontStyle);
                if (style == 2) // 이탤릭
                {
                    run.FontStyle = FontStyles.Italic;
                }
            }
            catch
            {
                // 폰트 스타일 변환 실패 시 기본 스타일 사용
            }
        }
    }

    private void ApplyColorStyle(Run run, Dictionary<string, object> styleAttributes)
    {
        // 전경색 적용
        if (styleAttributes.TryGetValue("ForegroundColor", out var foregroundColor) && foregroundColor != null)
        {
            try
            {
                var color = Color.FromRgb(
                    (byte)((int)foregroundColor & 0xFF),
                    (byte)((int)foregroundColor >> 8 & 0xFF),
                    (byte)((int)foregroundColor >> 16 & 0xFF)
                );
                run.Foreground = new SolidColorBrush(color);
            }
            catch
            {
                // 색상 변환 실패 시 기본 색상 사용
            }
        }

        // 배경색 적용
        if (styleAttributes.TryGetValue("BackgroundColor", out var backgroundColor) && backgroundColor != null)
        {
            try
            {
                var color = Color.FromRgb(
                    (byte)((int)backgroundColor & 0xFF),
                    (byte)((int)backgroundColor >> 8 & 0xFF),
                    (byte)((int)backgroundColor >> 16 & 0xFF)
                );
                run.Background = new SolidColorBrush(color);
            }
            catch
            {
                // 색상 변환 실패 시 기본 배경색 사용
            }
        }
    }

    private void ApplyTextDecoration(Run run, Dictionary<string, object> styleAttributes)
    {
        if (styleAttributes.TryGetValue("UnderlineStyle", out var underlineStyle) && underlineStyle != null)
        {
            try
            {
                var style = (TextDecorationCollection)underlineStyle;
                if (style != null && style.Count > 0)
                {
                    run.TextDecorations = TextDecorations.Underline;
                }
            }
            catch
            {
                // 밑줄 스타일 변환 실패 시 기본 스타일 사용
            }
        }
    }

    public void Log(string message)
    {
        Dispatcher.Invoke(() =>
        {
            var paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run(message + "\n"));
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        });
    }

    public async Task LogAsync(string message)
    {
        await _logSemaphore.WaitAsync();
        try
        {
            await Dispatcher.InvokeAsync(() =>
            {
                var paragraph = new Paragraph();
                paragraph.Inlines.Add(new Run(message + "\n"));
                LogRichTextBox.Document.Blocks.Add(paragraph);
                LogRichTextBox.ScrollToEnd();
            });
        }
        finally
        {
            _logSemaphore.Release();
        }
    }

    public void LogWithStyle(string message, Dictionary<string, object> styleAttributes)
    {
        Dispatcher.Invoke(() =>
        {
            var paragraph = new Paragraph();
            var run = new Run(message + "\n");

            ApplyFontStyle(run, styleAttributes);
            ApplyColorStyle(run, styleAttributes);
            ApplyTextDecoration(run, styleAttributes);

            paragraph.Inlines.Add(run);
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        });
    }

    public void LogWithStylePerChar(string message, List<Dictionary<string, object>> styleAttributesList)
    {
        Dispatcher.Invoke(() =>
        {
            var paragraph = new Paragraph();
            
            for (int i = 0; i < message.Length; i++)
            {
                var styleAttributes = styleAttributesList[i];
                var run = new Run(message[i].ToString());

                ApplyFontStyle(run, styleAttributes);
                ApplyColorStyle(run, styleAttributes);
                ApplyTextDecoration(run, styleAttributes);

                paragraph.Inlines.Add(run);
            }

            // 줄바꿈 추가
            paragraph.Inlines.Add(new Run("\n"));
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        });
    }

    private async Task SendTestMessageButton_ClickAsync(object sender, RoutedEventArgs e)
    {
        try
        {
            string parameterText = ParameterTextBox.Text.Trim();
            string targetServer = (ServerComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Flask";
            Log($"선택된 서버: {targetServer}");
            Log($"Flask 클라이언트 상태: {(_flaskClient != null ? "초기화됨" : "null")}");

            if (string.IsNullOrWhiteSpace(parameterText))
            {
                Log("Parameter를 입력해주세요.");
                return;
            }

            // Parameter를 JSON으로 파싱
            JsonElement parameters;
            try
            {
                parameters = JsonSerializer.Deserialize<JsonElement>(parameterText);
            }
            catch (Exception ex)
            {
                Log($"Parameter JSON 파싱 실패: {ex.Message}");
                return;
            }

            // command가 파라미터에 포함되어 있는지 확인
            if (!parameters.TryGetProperty("command", out var commandElement))
            {
                Log("Parameter에 'command' 필드가 없습니다.");
                return;
            }

            string command = commandElement.GetString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(command))
            {
                Log("Command가 비어있습니다.");
                return;
            }

            // 전송할 메시지 형식 로그 출력
            string messageJson = JsonSerializer.Serialize(parameters, new JsonSerializerOptions { WriteIndented = true });
            Log($"전송할 메시지 형식:\n{messageJson}");

            // 선택된 서버에 메시지 전송
            if (targetServer == "Flask" && _flaskClient != null)
            {
                Log("Flask 서버로 메시지 전송 시작...");
                try 
                {
                    await _flaskClient.SendMessageAsync(parameters);
                    Log("Flask 서버로 메시지 전송 완료");
                }
                catch (Exception ex)
                {
                    Log($"Flask 서버로 메시지 전송 중 오류 발생: {ex.Message}");
                    throw;
                }
            }
            else if (targetServer == "Vue" && _vueHost != null)
            {
                Log("Vue 서버로 메시지 전송 시작...");
                await _vueHost.SendMessageToAll(parameters);
                Log("Vue 서버로 메시지 전송 완료");
            }
            
            // 입력 필드 초기화
            ParameterTextBox.Clear();
        }
        catch (Exception ex)
        {
            Log($"메시지 전송 중 오류 발생: {ex.Message}");
        }
    }

    private void SendTestMessageButton_Click(object sender, RoutedEventArgs e)
    {
        _ = SendTestMessageButton_ClickAsync(sender, e);
    }

    private void ShowAllChatsButton_Click(object sender, RoutedEventArgs e)
    {
        var chatDataList = ChatDataManager.Instance.GetAllChatData();
        
        if (chatDataList.Count == 0)
        {
            Log("등록된 채팅 데이터가 없습니다.");
            return;
        }

        Log($"총 {chatDataList.Count}개의 채팅 데이터가 있습니다:");
        
        foreach (var chatData in chatDataList)
        {
            // 구분선 추가
            var separatorStyle = new Dictionary<string, object>
            {
                { "FontWeight", 700.0 },
                { "ForegroundColor", 0x808080 }  // 회색
            };
            LogWithStyle("----------------------------------------", separatorStyle);

            // 채팅 ID 표시
            var headerStyle = new Dictionary<string, object>
            {
                { "FontWeight", 700.0 },
                { "ForegroundColor", 0x0000FF }  // 파란색
            };
            LogWithStyle($"채팅 ID: {chatData.ChatId}", headerStyle);

            // 생성 시간 표시
            var timestampStyle = new Dictionary<string, object>
            {
                { "FontWeight", 500.0 },
                { "ForegroundColor", 0x008000 }  // 초록색
            };
            LogWithStyle($"생성 시간: {chatData.GeneratedTimestamp}", timestampStyle);

            // 현재 프로그램 정보
            var programStyle = new Dictionary<string, object>
            {
                { "FontWeight", 600.0 },
                { "ForegroundColor", 0x800080 }  // 보라색
            };
            LogWithStyle("현재 프로그램:", programStyle);
            Log($"  - 파일 ID: {chatData.CurrentProgram.FileId}");
            Log($"  - 파일 타입: {chatData.CurrentProgram.FileType}");
            Log($"  - 파일명: {chatData.CurrentProgram.FileName}");
            Log($"  - 파일 경로: {chatData.CurrentProgram.FilePath}");
            Log($"  - 컨텍스트: {chatData.CurrentProgram.Context}");
            Log($"  - 위치: {chatData.CurrentProgram.Position}");

            // 대상 프로그램 정보 (있는 경우)
            if (chatData.TargetProgram != null)
            {
                LogWithStyle("대상 프로그램:", programStyle);
                Log($"  - 파일 ID: {chatData.TargetProgram.FileId}");
                Log($"  - 파일 타입: {chatData.TargetProgram.FileType}");
                Log($"  - 파일명: {chatData.TargetProgram.FileName}");
                Log($"  - 파일 경로: {chatData.TargetProgram.FilePath}");
                Log($"  - 컨텍스트: {chatData.TargetProgram.Context}");
            }

            // 텍스트 내용
            var textStyle = new Dictionary<string, object>
            {
                { "FontWeight", 600.0 },
                { "ForegroundColor", 0xFF4500 }  // 주황색
            };
            LogWithStyle($"텍스트 내용 ({chatData.Texts.Count}개):", textStyle);
            foreach (var text in chatData.Texts)
            {
                Log($"  - 타입: {text.Type}");
                Log($"  - 내용: {text.Content}");
            }
        }
    }

    private async void ApplyContextButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // 테스트용 데이터 생성
            var testContext = new
            {
                Text = "<span style='font-family: 맑은 고딕; font-size: 12pt; color: #0000FF; background-color: #FFFF00'><b>테스트 텍스트입니다.</b></span>",
                Position = "1-1" // Word의 경우 "1-1", Excel의 경우 "R1C1-R1C1"
            };

            // 테스트용 프로그램 정보 생성
            var testProgram = new
            {
                FileType = "Word", // 또는 "Excel"
                FilePath = @"C:\Users\beste\OneDrive\Desktop\testFolder\single_test.docx" // 실제 테스트할 파일 경로
            };

            // Writer 생성 및 적용
            var writer = ContextWriterFactory.CreateWriter(testProgram.FileType);
            if (writer == null)
            {
                Log("지원하지 않는 프로그램입니다.");
                return;
            }

            // 파일 열기
            Log("파일 열기 시도...");
            if (!writer.OpenFile(testProgram.FilePath))
            {
                Log("파일을 열 수 없습니다.");
                return;
            }
            Log("파일 열기 성공");

            // 컨텍스트 적용
            Log("컨텍스트 적용 시도...");
            bool success = writer.ApplyTextWithStyle(
                testContext.Text,
                testContext.Position
            );

            if (success)
            {
                Log("컨텍스트 적용 성공");
            }
            else
            {
                Log("컨텍스트 적용 실패");
            }
        }
        catch (Exception ex)
        {
            Log($"오류 발생: {ex.Message}");
        }
    }

    private async void ApplyExcelContextButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // 테스트용 데이터 생성
            var testContext = new
            {
                Text = @"<table style='border-collapse: collapse;'><tr><td style='background-color: #FFFF00; color: #0000FF; font-size: 12pt'><b>테스트 셀 1</b></td><td style='background-color: #E6E6FA'>테스트 셀 2</td><td style=''></td></tr><tr><td style='background-color: #E6E6FA'>테스트 셀 3</td><td style='background-color: #FFFF00; color: #0000FF; font-size: 12pt'><b>테스트 셀 4</b></td><td style=''>e</td></tr><tr><td style=''>segseg</td><td style=''>g</td><td style=''>ge</td></tr><tr><td style=''>segseg</td><td style=''></td><td style=''></td></tr><tr><td style=''>eg</td><td style=''></td><td style=''>gdas</td></tr></table>",
                Position = "R1C1-R2C2" // Excel의 경우 "R1C1-R2C2" 형식 사용
            };

            // 테스트용 프로그램 정보 생성
            var testProgram = new
            {
                FileType = "Excel",
                FilePath = @"C:\Users\beste\OneDrive\Desktop\testFolder\single_test.xlsx" // 실제 테스트할 파일 경로
            };

            // Writer 생성 및 적용
            var writer = ContextWriterFactory.CreateWriter(testProgram.FileType);
            if (writer == null)
            {
                Log("지원하지 않는 프로그램입니다.");
                return;
            }

            // 파일 열기
            Log("파일 열기 시도...");
            if (!writer.OpenFile(testProgram.FilePath))
            {
                Log("파일을 열 수 없습니다.");
                return;
            }
            Log("파일 열기 성공");

            // 컨텍스트 적용
            Log("컨텍스트 적용 시도...");
            bool success = writer.ApplyTextWithStyle(
                testContext.Text,
                testContext.Position
            );

            if (success)
            {
                Log("컨텍스트 적용 성공");
            }
            else
            {
                Log("컨텍스트 적용 실패");
            }
        }
        catch (Exception ex)
        {
            Log($"오류 발생: {ex.Message}");
        }
    }

    private async void ApplyPPTContextButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string contextText = PPTContextTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(contextText))
            {
                Log("PPT 컨텍스트를 입력해주세요.");
                return;
            }

            // 테스트용 데이터 생성
            var testContext = new
            {
                Text = contextText,
                Position = "Slide 1"
            };

            // 테스트용 프로그램 정보 생성
            var testProgram = new
            {
                FileType = "PowerPoint",
                FilePath = @"C:\Users\beste\OneDrive\Desktop\testData\single_test.pptx" // 실제 테스트할 파일 경로
            };

            // Writer 생성 및 적용
            var writer = ContextWriterFactory.CreateWriter(testProgram.FileType);
            if (writer == null)
            {
                Log("지원하지 않는 프로그램입니다.");
                return;
            }

            // 파일 열기
            Log("파일 열기 시도...");
            if (!writer.OpenFile(testProgram.FilePath))
            {
                Log("파일을 열 수 없습니다.");
                return;
            }
            Log("파일 열기 성공");

            // 컨텍스트 적용
            Log("컨텍스트 적용 시도...");
            bool success = writer.ApplyTextWithStyle(
                testContext.Text,
                testContext.Position
            );

            if (success)
            {
                Log("컨텍스트 적용 성공");
            }
            else
            {
                Log("컨텍스트 적용 실패");
            }
        }
        catch (Exception ex)
        {
            Log($"오류 발생: {ex.Message}");
        }
    }

    private void ParameterTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ParameterTypeComboBox.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content != null && ParameterTextBox != null)
        {
            string template = selectedItem.Content.ToString() switch
            {
                "DisplayTextMessage" => JsonSerializer.Serialize(new overlay_gpt.Network.Models.Vue.DisplayText
                {
                    Command = "display_text",
                    ChatId = 1,
                    CurrentProgram = new ProgramInfo
                    {
                        Context = "샘플 컨텍스트"
                    },
                    TargetProgram = null,
                    Texts = new List<TextData>
                    {
                        new TextData
                        {
                            Type = "text_plain",
                            Content = "일반 텍스트 메시지입니다."
                        },
                        new TextData
                        {
                            Type = "text_block",
                            Content = "<b><font size=\"22\">보고서</font></b> <br> <p>샘플 내용입니다.</p>"
                        },
                        new TextData
                        {
                            Type = "table_block",
                            Content = JsonSerializer.Serialize(new List<List<string>>
                            {
                                new List<string> { "<b><color=\"blue\">제목</color></b>" },
                                new List<string> { "내용1" },
                                new List<string> { "내용2" }
                            })
                        },
                        new TextData
                        {
                            Type = "code_block",
                            Content = "int a = 0;\nConsole.WriteLine(a);"
                        }
                    }
                }, new JsonSerializerOptions { WriteIndented = true }),
                "ProgramInfo" => JsonSerializer.Serialize(new
                {
                    id = 1,
                    type = "C#",
                    context = "샘플 컨텍스트"
                }, new JsonSerializerOptions { WriteIndented = true }),
                "Custom JSON" => "{\n    // 여기에 커스텀 JSON을 입력하세요\n}",
                _ => "{}"
            };

            ParameterTextBox.Text = template;
        }
    }

    private async void TestNtfsFileFinderButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // 테스트용 fileId와 volumeId
            long fileId = 21673573207065453; // 실제 테스트할 파일의 ID
            uint volumeId = 2524257335; // 실제 테스트할 볼륨의 ID
            
            Log("NtfsFileFinder 테스트 시작...");
            Log($"테스트할 FileId: {fileId}");
            Log($"테스트할 VolumeId: {volumeId}");

            // fileId와 volumeId로 파일 찾기
            string foundPath = _ntfsFileFinder.FindFileByFileIdAndVolumeId(fileId, volumeId);
            
            if (foundPath != null)
            {
                Log($"찾은 파일 경로: {foundPath}");
                Log("테스트 성공!");
            }
            else
            {
                Log("파일을 찾을 수 없습니다.");
            }
        }
        catch (Exception ex)
        {
            Log($"테스트 중 오류 발생: {ex.Message}");
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        
        // 서버 연결 종료
        if (_flaskClient != null)
        {
            _ = _flaskClient.DisconnectAsync();
        }
        
        if (_vueHost != null)
        {
            _ = _vueHost.StopAsync();
        }
    }
} 