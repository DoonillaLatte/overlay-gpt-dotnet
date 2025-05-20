using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Controls;
using System.Text.Json;
using System.Threading;

namespace overlay_gpt;

public partial class LogWindow : Window
{
    private static LogWindow? _instance;
    private static readonly SemaphoreSlim _instanceSemaphore = new SemaphoreSlim(1, 1);
    private static readonly SemaphoreSlim _logSemaphore = new SemaphoreSlim(1, 1);

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

    private LogWindow()
    {
        InitializeComponent();
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
            string command = CommandTextBox.Text.Trim();
            string parameterText = ParameterTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(command))
            {
                Log("Command를 입력해주세요.");
                return;
            }

            if (string.IsNullOrWhiteSpace(parameterText))
            {
                Log("Parameter를 입력해주세요.");
                return;
            }

            // Parameter를 JSON으로 파싱
            object? parameters;
            try
            {
                parameters = JsonSerializer.Deserialize<object>(parameterText);
            }
            catch (Exception ex)
            {
                Log($"Parameter JSON 파싱 실패: {ex.Message}");
                return;
            }

            // 전송할 메시지 형식 생성
            var message = new
            {
                command = command,
                parameters = parameters
            };

            // 전송할 메시지 형식 로그 출력
            string messageJson = JsonSerializer.Serialize(message, new JsonSerializerOptions { WriteIndented = true });
            Log($"전송할 메시지 형식:\n{messageJson}");
            
            // 입력 필드 초기화
            CommandTextBox.Clear();
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
} 