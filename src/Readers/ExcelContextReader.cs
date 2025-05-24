using System.Collections.Generic;
using System.Windows.Automation;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private Excel.Application excelApp;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        [return: MarshalAs(UnmanagedType.IUnknown)]
        private static extern object CLSIDFromString(string lpsz, out Guid pclsid);

        [DllImport("ole32.dll", PreserveSig = false)]
        [return: MarshalAs(UnmanagedType.IUnknown)]
        private static extern object GetActiveObject([In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved);

        private static object GetActiveCOMObject(string progID)
        {
            Console.WriteLine($"GetActiveCOMObject 호출됨: {progID}");
            try
            {
                Guid clsid;
                CLSIDFromString(progID, out clsid);
                object obj = GetActiveObject(clsid, IntPtr.Zero);
                Console.WriteLine("GetActiveObject 성공");
                return obj;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetActiveObject 실패: {ex.Message}");
                throw;
            }
        }

        public ExcelContextReader()
        {
            Console.WriteLine("ExcelContextReader 생성자 시작");
            try
            {
                Console.WriteLine("기존 Excel 인스턴스 찾기 시도");
                excelApp = (Excel.Application)GetActiveCOMObject("Excel.Application");
                Console.WriteLine("기존 Excel 인스턴스 찾음");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 인스턴스 찾기 실패: {ex.Message}");
                throw new System.Exception("실행 중인 Excel을 찾을 수 없습니다.");
            }
            Console.WriteLine("ExcelContextReader 생성자 완료");
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            Console.WriteLine("GetSelectedTextWithStyle 메서드 시작");
            if (excelApp == null || excelApp.Selection == null)
            {
                Console.WriteLine("Excel 앱 또는 선택 영역이 null입니다");
                return (string.Empty, new Dictionary<string, object>());
            }

            try
            {
                Excel.Range selection = excelApp.Selection;
                string selectedText = selection.Text.ToString();
                Console.WriteLine($"선택된 텍스트: {selectedText}");

                var styleAttributes = new Dictionary<string, object>
                {
                    ["FontName"] = selection.Font.Name,
                    ["FontSize"] = selection.Font.Size,
                    ["FontWeight"] = selection.Font.Bold ? 700 : 400,
                    ["ForegroundColor"] = selection.Font.Color,
                    ["BackgroundColor"] = selection.Interior.Color,
                    ["UnderlineStyle"] = selection.Font.Underline ? 1 : 0
                };

                Console.WriteLine($"스타일 속성 추출 완료: FontName={styleAttributes["FontName"]}, FontSize={styleAttributes["FontSize"]}");
                return (selectedText, styleAttributes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 또는 스타일 추출 중 오류 발생: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>());
            }
        }
    }
}
