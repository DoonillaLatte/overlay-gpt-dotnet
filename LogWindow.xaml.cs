using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Controls;

namespace overlay_gpt;

public partial class LogWindow : Window
{
    private static LogWindow? _instance;

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
            var run = new Run(message + "\n");
            paragraph.Inlines.Add(run);
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        });
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
} 