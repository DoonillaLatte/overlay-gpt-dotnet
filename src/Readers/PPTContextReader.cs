using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using HtmlAgilityPack;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPTApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace overlay_gpt
{
    public class PPTContextReader : BaseContextReader
    {
        private PPTApp? _pptApp;
        private Presentation? _presentation;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        private static object GetActiveObject(string progID)
        {
            CLSIDFromProgID(progID, out Guid clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            return obj!;
        }

        /// <summary>
        /// 도형(shape) 또는 텍스트가 선택된 상태에서 Copy()를 호출하면
        /// 클립보드에 올라가는 HTML(Fragment) 또는 GVML 패키지(Zip)에서
        /// 실제 도형 정보를 담고 있는 "clipboard/drawings/drawing1.xml"만 꺼내 반환합니다.
        /// </summary>
        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                Console.WriteLine("PowerPoint 데이터 읽기 시작...");

                // 1) 실행 중인 PowerPoint 프로세스 찾기
                var pptProcesses = Process.GetProcessesByName("POWERPNT");
                if (pptProcesses.Length == 0)
                    throw new InvalidOperationException("PowerPoint가 실행 중이지 않습니다.");

                // 2) 포그라운드(활성) PowerPoint 프로세스 찾기
                Process? activePPTProcess = null;
                foreach (var proc in pptProcesses)
                {
                    if (proc.MainWindowHandle != IntPtr.Zero &&
                        !string.IsNullOrEmpty(proc.MainWindowTitle) &&
                        proc.MainWindowHandle == GetForegroundWindow())
                    {
                        activePPTProcess = proc;
                        break;
                    }
                }
                if (activePPTProcess == null)
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);

                // 3) COM으로 PowerPoint.Application 객체 가져오기
                _pptApp = (PPTApp)GetActiveObject("PowerPoint.Application");
                _presentation = _pptApp.ActivePresentation;
                if (_presentation == null)
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);

                Console.WriteLine($"- 프레젠테이션: {_presentation.Name}");
                Console.WriteLine($"- 경로: {_presentation.FullName}");
                Console.WriteLine($"- 저장 여부: {(_presentation.Saved == MsoTriState.msoTrue ? "저장됨" : "저장되지 않음")}");

                // 4) 현재 Selection 가져오기
                var selection = _pptApp.ActiveWindow.Selection;
                if (selection == null)
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);

                Console.WriteLine($"Selection.Type = {selection.Type}");

                // 5) 클립보드 초기화
                Clipboard.Clear();

                // 6) 선택된 항목에 따라 Copy 호출
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    // 도형 선택 시 selection.Copy()로 GVML 패키지를 클립보드에 올려줌
                    try
                    {
                        selection.Copy();
                    }
                    catch (COMException)
                    {
                        // 만약 실패하면, ShapeRange.Copy()로 재시도
                        try { selection.ShapeRange.Copy(); }
                        catch (COMException ex)
                        {
                            Console.WriteLine($"도형 복사 실패: {ex.Message}");
                        }
                    }
                }
                else if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    // 텍스트 선택 시, HTML Fragment를 클립보드에 올려줌
                    try { selection.Copy(); }
                    catch (COMException ex)
                    {
                        Console.WriteLine($"텍스트 복사 중 오류: {ex.Message}");
                    }
                }
                else
                {
                    // 도형/텍스트 외 항목인 경우에도 일단 Copy() 시도
                    try { selection.Copy(); }
                    catch (COMException ex)
                    {
                        Console.WriteLine($"Copy() 호출 중 예외: {ex.Message}");
                    }
                }

                // 7) 클립보드 갱신 대기
                Thread.Sleep(200);

                // 8) 클립보드 데이터 가져오기
                IDataObject dataObj = Clipboard.GetDataObject()!;
                if (dataObj == null)
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);

                // 9) 텍스트(HTML Fragment) 우선 처리
                if (Clipboard.ContainsText(TextDataFormat.Html))
                {
                    string htmlRaw = Clipboard.GetText(TextDataFormat.Html)!;
                    return ProcessHtmlFragment(htmlRaw);
                }

                // 10) GVML 패키지(Zip) 처리: "Art::GVML ClipFormat" 포맷이 있는지 확인
                string? gvmlKey = dataObj.GetFormats()
                    .FirstOrDefault(f => f.Equals("Art::GVML ClipFormat", StringComparison.OrdinalIgnoreCase));

                if (gvmlKey != null)
                {
                    Console.WriteLine($"GVML 패키지 포맷 발견: \"{gvmlKey}\"");
                    object rawData = dataObj.GetData(gvmlKey)!;

                    // byte[] 또는 MemoryStream 형태로 올 수 있음
                    byte[] zipBytes;
                    if (rawData is byte[] ba)
                    {
                        zipBytes = ba;
                    }
                    else if (rawData is MemoryStream ms)
                    {
                        zipBytes = ms.ToArray();
                    }
                    else
                    {
                        Console.WriteLine($"예상치 못한 GVML 데이터 타입: {rawData.GetType().Name}");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    return ProcessGvmlZip(zipBytes);
                }

                // 11) 위 두 포맷이 모두 없으면 빈 문자열 반환
                Console.WriteLine("HTML/ GVML 관련 클립보드 포맷을 찾지 못했습니다.");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"PowerPoint COM 오류: {comEx.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            finally
            {
                if (_presentation != null)
                    Marshal.ReleaseComObject(_presentation);
                if (_pptApp != null)
                    Marshal.ReleaseComObject(_pptApp);
            }
        }

        /// <summary>
        /// HTML Fragment(rawHtml)에서 <!--StartFragment-->와 <!--EndFragment--> 사이만 추출하여
        /// test.html로 저장하고 fragment만 반환합니다.
        /// </summary>
        private (string, Dictionary<string, object>, string) ProcessHtmlFragment(string rawHtml)
        {
            int startIdx = rawHtml.IndexOf("<!--StartFragment-->");
            int endIdx = rawHtml.IndexOf("<!--EndFragment-->");
            if (startIdx != -1 && endIdx != -1 && endIdx > startIdx)
            {
                int fragContentStart = startIdx + "<!--StartFragment-->".Length;
                int fragLength = endIdx - fragContentStart;
                string fragment = rawHtml.Substring(fragContentStart, fragLength);

                Console.WriteLine("=== 추출된 HTML Fragment (일부) ===");
                Console.WriteLine(fragment.Substring(0, Math.Min(fragment.Length, 500)));
                if (fragment.Length > 500) Console.WriteLine("... (생략) ...");
                Console.WriteLine("=== 끝 ===");

                try
                {
                    string htmlTemplate = @"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Clipboard HTML Fragment (길이: {1}자)</title>
</head>
<body>
{0}
</body>
</html>";
                    string fullHtml = string.Format(htmlTemplate, fragment, fragment.Length);
                    File.WriteAllText("test.html", fullHtml, System.Text.Encoding.UTF8);
                    Console.WriteLine("▶ test.html 파일 저장 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"test.html 저장 오류: {ex.Message}");
                }

                return (fragment, new Dictionary<string, object>(), string.Empty);
            }
            else
            {
                Console.WriteLine("HTML Fragment 마커를 찾을 수 없습니다. 전체 HTML을 test_full.html로 저장합니다.");
                try
                {
                    File.WriteAllText("test_full.html", rawHtml, System.Text.Encoding.UTF8);
                    Console.WriteLine("▶ test_full.html 파일 저장 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"test_full.html 저장 오류: {ex.Message}");
                }

                return (rawHtml, new Dictionary<string, object>(), string.Empty);
            }
        }

        /// <summary>
        /// GVML 패키지(zipBytes)에서 "clipboard/drawings/drawing1.xml"만 추출하여
        /// test_drawing1.xml로 저장하고, 그 내용 문자열만 반환합니다.
        /// </summary>
        private (string, Dictionary<string, object>, string) ProcessGvmlZip(byte[] zipBytes)
        {
            try
            {
                using var mem = new MemoryStream(zipBytes);
                using var archive = new ZipArchive(mem, ZipArchiveMode.Read, leaveOpen: true);

                // "clipboard/drawings/drawing1.xml" 경로를 가진 엔트리를 찾음
                // (실제 ZIP 내부 구조에 따라 경로가 달라질 수 있음)
                var entry = archive.GetEntry("clipboard/drawings/drawing1.xml")
                         ?? archive.Entries.FirstOrDefault(e => e.FullName.EndsWith("drawing1.xml", StringComparison.OrdinalIgnoreCase));

                if (entry == null)
                {
                    Console.WriteLine("ZIP 내부에서 drawing1.xml 엔트리를 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // 해당 엔트리의 내용을 문자열로 읽어들임
                string xmlContent;
                using (var reader = new StreamReader(entry.Open(), System.Text.Encoding.UTF8))
                {
                    xmlContent = reader.ReadToEnd();
                }

                Console.WriteLine("=== 추출된 drawing1.xml 내용 (일부) ===");
                Console.WriteLine(xmlContent.Substring(0, Math.Min(xmlContent.Length, 500)));
                if (xmlContent.Length > 500) Console.WriteLine("... (생략) ...");
                Console.WriteLine("=== 끝 ===");

                // 파일로 저장
                try
                {
                    File.WriteAllText("test_drawing1.xml", xmlContent, System.Text.Encoding.UTF8);
                    Console.WriteLine("▶ test_drawing1.xml 파일 저장 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"test_drawing1.xml 저장 오류: {ex.Message}");
                }

                // 최종 반환값: XML 문자열
                return (xmlContent, new Dictionary<string, object>(), string.Empty);
            }
            catch (InvalidDataException)
            {
                Console.WriteLine("GVML 패키지를 ZIP으로 열 수 없습니다. (잘못된 형식)");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GVML ZIP 처리 중 오류: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
        }

        // 이하 기존 헬퍼 메서드는 그대로 두거나 필요에 따라 유지하세요.
        private string CleanAndNormalizeHtml(string rawFragment)
        {
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml("<div id=\"wrapper\">" + rawFragment + "</div>");
            HtmlNode wrapper = htmlDoc.GetElementbyId("wrapper")!;

            RemoveUnwantedNodes(wrapper);
            RemoveWordSpecificAttributes(wrapper);
            RemoveClassAttributes(wrapper);
            RemoveEmptySpans(wrapper);
            MergeAdjacentSpans(wrapper);
            RemoveEmptyStyleAttributes(wrapper);
            ProcessImages(wrapper);

            string interimHtml = wrapper.InnerHtml;
            return NormalizeWhitespace(interimHtml);
        }

        private void RemoveUnwantedNodes(HtmlNode root)
        {
            var metas = root.SelectNodes("//meta");
            if (metas != null) foreach (var meta in metas) meta.Remove();

            var xmlNodes = root.SelectNodes("//xml");
            if (xmlNodes != null) foreach (var node in xmlNodes) node.Remove();

            var allNodes = root.SelectNodes("//*");
            if (allNodes != null)
            {
                foreach (var node in allNodes.ToList())
                {
                    if (node.Name.StartsWith("o:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("v:", StringComparison.OrdinalIgnoreCase))
                    {
                        node.Remove();
                    }
                }
            }

            var comments = root.SelectNodes("//comment()");
            if (comments != null) foreach (var c in comments.Cast<HtmlCommentNode>()) c.Remove();
        }

        private void RemoveWordSpecificAttributes(HtmlNode root)
        {
            var nodesWithLang = root.SelectNodes("//*[@lang]");
            if (nodesWithLang != null) foreach (var node in nodesWithLang) node.Attributes.Remove("lang");

            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var styleAttr = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(styleAttr))
                    {
                        node.Attributes.Remove("style");
                        continue;
                    }

                    var declarations = styleAttr
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(d => d.Trim())
                        .Where(d => !d.StartsWith("mso-", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (declarations.Any())
                        node.SetAttributeValue("style", string.Join(";", declarations) + ";");
                    else
                        node.Attributes.Remove("style");
                }
            }
        }

        private void RemoveClassAttributes(HtmlNode root)
        {
            var nodesWithClass = root.SelectNodes("//*[@class]");
            if (nodesWithClass != null) foreach (var node in nodesWithClass) node.Attributes.Remove("class");
        }

        private void RemoveEmptySpans(HtmlNode root)
        {
            var spans = root.SelectNodes("//span");
            if (spans != null)
            {
                foreach (var span in spans.ToList())
                {
                    string styleAttr = span.GetAttributeValue("style", "").Trim();
                    string inner = span.InnerHtml.Trim();
                    if (string.IsNullOrEmpty(styleAttr) && string.IsNullOrEmpty(inner))
                        span.Remove();
                }
            }
        }

        private void MergeAdjacentSpans(HtmlNode root)
        {
            var parentNodes = root.SelectNodes("//*");
            if (parentNodes == null) return;

            foreach (var parent in parentNodes)
            {
                var children = parent.ChildNodes.ToList();
                for (int i = 0; i < children.Count - 1; i++)
                {
                    var curr = children[i];
                    var next = children[i + 1];
                    if (curr.Name.Equals("span", StringComparison.OrdinalIgnoreCase)
                        && next.Name.Equals("span", StringComparison.OrdinalIgnoreCase))
                    {
                        string styleCurr = curr.GetAttributeValue("style", "");
                        string styleNext = next.GetAttributeValue("style", "");
                        if (styleCurr == styleNext)
                        {
                            curr.InnerHtml += next.InnerHtml;
                            next.Remove();
                            children = parent.ChildNodes.ToList();
                            i--;
                        }
                    }
                }
            }
        }

        private void RemoveEmptyStyleAttributes(HtmlNode root)
        {
            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var val = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(val))
                        node.Attributes.Remove("style");
                }
            }
        }

        private void ProcessImages(HtmlNode root)
        {
            string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
            if (!Directory.Exists(imageDir))
                Directory.CreateDirectory(imageDir);

            var images = root.SelectNodes("//img");
            if (images != null)
            {
                foreach (var img in images)
                {
                    string src = img.GetAttributeValue("src", "");
                    if (src.StartsWith("data:image"))
                    {
                        try
                        {
                            string[] parts = src.Split(',');
                            if (parts.Length > 1)
                            {
                                string imageData = parts[1];
                                string imageId = Guid.NewGuid().ToString();
                                string imagePath = Path.Combine(imageDir, $"{imageId}.jpg");
                                byte[] imageBytes = Convert.FromBase64String(imageData);
                                File.WriteAllBytes(imagePath, imageBytes);
                                img.SetAttributeValue("src", $"images/{imageId}.jpg");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"이미지 처리 오류: {ex.Message}");
                        }
                    }
                }
            }
        }

        private string NormalizeWhitespace(string html)
        {
            return string.IsNullOrEmpty(html) ? string.Empty : html.Trim();
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            PPTApp? tempPPTApp = null;
            Presentation? tempPresentation = null;

            try
            {
                tempPPTApp = (PPTApp)GetActiveObject("PowerPoint.Application");
                tempPresentation = tempPPTApp.ActivePresentation;
                if (tempPresentation == null)
                    return (null, null, "PowerPoint", string.Empty, string.Empty);

                string filePath = tempPresentation.FullName;
                string fileName = tempPresentation.Name;
                return (null, null, "PowerPoint", fileName, filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "PowerPoint", string.Empty, string.Empty);
            }
            finally
            {
                if (tempPresentation != null) Marshal.ReleaseComObject(tempPresentation);
                if (tempPPTApp != null)     Marshal.ReleaseComObject(tempPPTApp);
            }
        }
    }
}
