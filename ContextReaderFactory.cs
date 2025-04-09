using System.Windows.Automation;

namespace overlay_gpt
{
    public static class ContextReaderFactory
    {
        public static IContextReader CreateReader(AutomationElement element)
        {
            if (element == null)
                return new TextPatternContextReader();

            if (element.TryGetCurrentPattern(TextPattern.Pattern, out _))
                return new TextPatternContextReader();
            
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out _))
                return new ValuePatternContextReader();

            return new TextPatternContextReader();
        }
    }
} 