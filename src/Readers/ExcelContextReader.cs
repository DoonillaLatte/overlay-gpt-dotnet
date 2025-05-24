using System.Collections.Generic;
using System.Windows.Automation;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private Excel.Application excelApp;

        public ExcelContextReader()
        {
            try
            {
                excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                try
                {
                    excelApp = new Excel.Application();
                }
                catch
                {
                    throw new System.Exception("Excel이 실행 중이 아닙니다.");
                }
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            if (excelApp == null || excelApp.Selection == null)
                return (string.Empty, new Dictionary<string, object>());

            Excel.Range selection = excelApp.Selection;
            string selectedText = selection.Text.ToString();

            var styleAttributes = new Dictionary<string, object>
            {
                ["FontName"] = selection.Font.Name,
                ["FontSize"] = selection.Font.Size,
                ["FontWeight"] = selection.Font.Bold ? 700 : 400,
                ["ForegroundColor"] = selection.Font.Color,
                ["BackgroundColor"] = selection.Interior.Color,
                ["UnderlineStyle"] = selection.Font.Underline ? 1 : 0
            };

            return (selectedText, styleAttributes);
        }
    }
}
