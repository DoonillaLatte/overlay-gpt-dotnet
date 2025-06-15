using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Automation.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Xml;

namespace overlay_gpt
{
    public class NewWordContextReader : BaseContextReader
    {
        private Word.Application _wordApp;
        private bool _isTargetProg;
        private string _filePath;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public uint dwVolumeSerialNumber;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint nNumberOfLinks;
            public uint nFileIndexHigh;
            public uint nFileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILETIME
        {
            public uint dwLowDateTime;
            public uint dwHighDateTime;
        }

        private const uint GENERIC_READ = 0x80000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        private (ulong FileId, uint VolumeId)? GetFileId(string filePath)
        {
            try
            {
                Console.WriteLine($"GetFileId 호출 - 파일 경로: {filePath}");
                
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("파일이 존재하지 않습니다.");
                    return null;
                }

                IntPtr handle = CreateFile(
                    filePath,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero);

                if (handle.ToInt64() == -1)
                {
                    Console.WriteLine($"CreateFile 실패 - 에러 코드: {Marshal.GetLastWin32Error()}");
                    return null;
                }

                try
                {
                    BY_HANDLE_FILE_INFORMATION fileInfo;
                    if (GetFileInformationByHandle(handle, out fileInfo))
                    {
                        ulong fileId = ((ulong)fileInfo.nFileIndexHigh << 32) | fileInfo.nFileIndexLow;
                        Console.WriteLine($"파일 ID 정보 가져오기 성공:");
                        Console.WriteLine($"- FileId: {fileId}");
                        Console.WriteLine($"- VolumeId: {fileInfo.dwVolumeSerialNumber}");
                        return (fileId, fileInfo.dwVolumeSerialNumber);
                    }
                    else
                    {
                        Console.WriteLine($"GetFileInformationByHandle 실패 - 에러 코드: {Marshal.GetLastWin32Error()}");
                    }
                }
                finally
                {
                    CloseHandle(handle);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 ID 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            }
            return null;
        }

        public NewWordContextReader(bool isTargetProg = false, string filePath = "")
        {
            Console.WriteLine($"WordContextReader 생성 시도 - isTargetProg: {isTargetProg}");
            _isTargetProg = isTargetProg;
            _filePath = filePath;
        }

        private bool IsWordProcessActive()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                if (foregroundWindow == IntPtr.Zero)
                {
                    Console.WriteLine("포커스된 창을 찾을 수 없습니다.");
                    return false;
                }

                uint processId;
                GetWindowThreadProcessId(foregroundWindow, out processId);

                Process foregroundProcess = Process.GetProcessById((int)processId);
                Console.WriteLine($"현재 포커스된 프로세스: {foregroundProcess.ProcessName} (PID: {processId})");

                return foregroundProcess.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"프로세스 확인 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber)
            GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                Console.WriteLine("Word 데이터 읽기 시작...");
                
                _wordApp = (Word.Application)GetActiveObject("Word.Application");
                
                Console.WriteLine($"Word 애플리케이션 상태: {(_wordApp == null ? "null" : "존재함")}");

                if (_wordApp == null)
                {
                    Console.WriteLine("Word 애플리케이션이 null입니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                if (!IsWordProcessActive())
                {
                    Console.WriteLine("현재 포커스된 프로세스가 Word가 아닙니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                try
                {
                    Console.WriteLine("활성 문서 가져오기 시도...");
                    if (_wordApp.Documents.Count == 0)
                    {
                        Console.WriteLine("열린 문서가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    var activeDoc = _wordApp.ActiveDocument;
                    if (activeDoc == null)
                    {
                        Console.WriteLine("활성 문서가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }
                    Console.WriteLine($"활성 문서: {activeDoc.Name}");

                    // 로그 윈도우의 파일 경로 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.FilePathTextBox.Text = activeDoc.FullName;
                    });

                    Console.WriteLine("Selection 객체 가져오기 시도...");
                    var sel = _wordApp.Selection;
                    Console.WriteLine($"Selection 객체 상태: {(sel == null ? "null" : "존재함")}");
                    
                    if (sel == null)
                    {
                        Console.WriteLine("Selection 객체가 null입니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // 선택된 텍스트의 위치 정보 가져오기
                    string position = $"시작: {sel.Start}, 끝: {sel.End}";
                    Console.WriteLine($"선택된 텍스트 위치: {position}");

                    // 로그 윈도우의 위치 정보 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.PositionTextBox.Text = position;
                    });

                    Console.WriteLine($"선택된 텍스트 길이: {sel.Text?.Length ?? 0}");
                    Console.WriteLine($"선택된 텍스트 내용: {sel.Text}");

                    if (string.IsNullOrEmpty(sel.Text))
                    {
                        Console.WriteLine("선택된 텍스트가 비어있습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // HTML 형식으로 클립보드 복사 시도
                    try
                    {
                        Console.WriteLine("클립보드 복사 시도...");
                        sel.Copy();
                        Console.WriteLine("클립보드 복사 완료");

                        Console.WriteLine("클립보드 형식 확인 중...");
                        Console.WriteLine($"HTML 형식 존재: {Clipboard.ContainsText(TextDataFormat.Html)}");
                        Console.WriteLine($"일반 텍스트 존재: {Clipboard.ContainsText()}");
                        Console.WriteLine($"RTF 형식 존재: {Clipboard.ContainsText(TextDataFormat.Rtf)}");

                        if (Clipboard.ContainsText(TextDataFormat.Html))
                        {
                            string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                            Console.WriteLine($"HTML 데이터 길이: {htmlContent.Length}");
                            Console.WriteLine("HTML 데이터 일부: " + htmlContent.Substring(0, Math.Min(100, htmlContent.Length)));
                            
                            // HTML 프래그먼트 추출
                            int startIdx = htmlContent.IndexOf("<!--StartFragment-->");
                            int endIdx = htmlContent.IndexOf("<!--EndFragment-->");
                            Console.WriteLine($"StartFragment 위치: {startIdx}");
                            Console.WriteLine($"EndFragment 위치: {endIdx}");

                            if (startIdx != -1 && endIdx != -1 && endIdx > startIdx)
                            {
                                int fragContentStart = startIdx + "<!--StartFragment-->".Length;
                                int fragLength = endIdx - fragContentStart;
                                string rawFragment = htmlContent.Substring(fragContentStart, fragLength);
                                Console.WriteLine($"추출된 HTML 프래그먼트 길이: {rawFragment.Length}");
                                
                                // 스타일 속성 가져오기
                                var font = sel.Range.Font;
                                Console.WriteLine("폰트 정보 가져오기:");
                                Console.WriteLine($"- 폰트 이름: {font.Name}");
                                Console.WriteLine($"- 폰트 크기: {font.Size}");
                                Console.WriteLine($"- 굵게: {font.Bold}");
                                Console.WriteLine($"- 기울임: {font.Italic}");
                                Console.WriteLine($"- 밑줄: {font.Underline}");
                                Console.WriteLine($"- 색상: {font.Color}");

                                var styleAttributes = new Dictionary<string, object>
                                {
                                    ["FontName"]        = font.Name,
                                    ["FontSize"]        = font.Size,
                                    ["FontBold"]        = font.Bold == 1,
                                    ["FontItalic"]      = font.Italic == 1,
                                    ["FontUnderline"]   = font.Underline.ToString(),
                                    ["FontColor"]       = font.Color.ToString()
                                };

                                var lineNumber = sel.Range.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                                Console.WriteLine($"라인 번호: {lineNumber}");

                                // 로그 윈도우의 컨텍스트 업데이트
                                LogWindow.Instance.Dispatcher.Invoke(() =>
                                {
                                    LogWindow.Instance.ContextTextBox.Text = rawFragment;
                                });

                                return (rawFragment, styleAttributes, position);
                            }
                            else
                            {
                                Console.WriteLine("HTML 프래그먼트 태그를 찾을 수 없습니다.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("클립보드에 HTML 형식 데이터가 없습니다.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"HTML 클립보드 복사 실패: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                        
                        // 수식 처리를 시도
                        Console.WriteLine("수식 처리 방식으로 전환...");
                        
                        // 스타일 속성 가져오기
                        var font = sel.Range.Font;
                        var styleAttributes = new Dictionary<string, object>
                        {
                            ["FontName"]        = font.Name,
                            ["FontSize"]        = font.Size,
                            ["FontBold"]        = font.Bold == 1,
                            ["FontItalic"]      = font.Italic == 1,
                            ["FontUnderline"]   = font.Underline.ToString(),
                            ["FontColor"]       = font.Color.ToString()
                        };

                        string processedText = ProcessSelectionWithEquations(sel);

                        // 로그 윈도우의 컨텍스트 업데이트
                        LogWindow.Instance.Dispatcher.Invoke(() =>
                        {
                            LogWindow.Instance.ContextTextBox.Text = processedText;
                        });

                        return (processedText, styleAttributes, position);
                    }

                    // HTML 형식이 실패한 경우 기본 텍스트 반환
                    Console.WriteLine("기본 텍스트 반환");

                    // 로그 윈도우의 컨텍스트 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.ContextTextBox.Text = sel.Text;
                    });

                    return (sel.Text, new Dictionary<string, object>(), position);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"문서 접근 중 오류 발생: {ex.Message}");
                    Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word 데이터 읽기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            Word.Application? tempWordApp = null;
            Word.Document? tempDocument = null;
            
            try
            {
                Console.WriteLine("Word COM 객체 가져오기 시도...");
                tempWordApp = (Word.Application)GetActiveObject("Word.Application");
                Console.WriteLine("Word COM 객체 가져오기 성공");

                if (_isTargetProg && !string.IsNullOrEmpty(_filePath))
                {
                    // 모든 문서 확인
                    foreach (Word.Document doc in tempWordApp.Documents)
                    {
                        try
                        {
                            string filePath = doc.FullName;
                            string fileName = doc.Name;
                            
                            Console.WriteLine($"Word 문서 정보:");
                            Console.WriteLine($"- 파일 경로: {filePath}");
                            Console.WriteLine($"- 파일 이름: {fileName}");
                            
                            if (string.IsNullOrEmpty(filePath))
                            {
                                Console.WriteLine("파일 경로가 비어있습니다.");
                                continue;
                            }
                            
                            if (string.Equals(_filePath, filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                var fileIdInfo = GetFileId(filePath);
                                
                                if (fileIdInfo == null)
                                {
                                    Console.WriteLine("파일 ID 정보를 가져오지 못했습니다.");
                                }
                                else
                                {
                                    Console.WriteLine($"파일 ID 정보:");
                                    Console.WriteLine($"- FileId: {fileIdInfo.Value.FileId}");
                                    Console.WriteLine($"- VolumeId: {fileIdInfo.Value.VolumeId}");
                                }
                                
                                return (
                                    fileIdInfo?.FileId,
                                    fileIdInfo?.VolumeId,
                                    "Word",
                                    fileName,
                                    filePath
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"문서 처리 중 오류: {ex.Message}");
                            continue;
                        }
                    }
                }
                else
                {
                    // 활성 문서 정보 가져오기
                    tempDocument = tempWordApp.ActiveDocument;
                    if (tempDocument != null)
                    {
                        string filePath = tempDocument.FullName;
                        string fileName = tempDocument.Name;
                        
                        Console.WriteLine($"활성 Word 문서 정보:");
                        Console.WriteLine($"- 파일 경로: {filePath}");
                        Console.WriteLine($"- 파일 이름: {fileName}");
                        
                        var fileIdInfo = GetFileId(filePath);
                        
                        if (fileIdInfo == null)
                        {
                            Console.WriteLine("파일 ID 정보를 가져오지 못했습니다.");
                        }
                        else
                        {
                            Console.WriteLine($"파일 ID 정보:");
                            Console.WriteLine($"- FileId: {fileIdInfo.Value.FileId}");
                            Console.WriteLine($"- VolumeId: {fileIdInfo.Value.VolumeId}");
                        }
                        
                        return (
                            fileIdInfo?.FileId,
                            fileIdInfo?.VolumeId,
                            "Word",
                            fileName,
                            filePath
                        );
                    }
                }
                
                Console.WriteLine("문서를 찾을 수 없습니다.");
                return (null, null, "Word", string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (null, null, "Word", string.Empty, string.Empty);
            }
            finally
            {
                if (tempDocument != null) Marshal.ReleaseComObject(tempDocument);
                if (tempWordApp != null) Marshal.ReleaseComObject(tempWordApp);
            }
        }

        private string ProcessSelectionWithEquations(Word.Selection selection)
        {
            var originalClipboard = Clipboard.GetDataObject();
            string result = string.Empty;

            try
            {
                Console.WriteLine("수식 처리 시작...");
                // 수식이 포함된 범위를 복사
                selection.Copy();
                Console.WriteLine("선택된 텍스트 복사 완료");

                if (Clipboard.ContainsData(DataFormats.Html))
                {
                    var cfHtml = Clipboard.GetData(DataFormats.Html) as string;
                    Console.WriteLine("HTML 데이터 발견");
                    result = ExtractHtmlFragment(cfHtml);
                    Console.WriteLine($"추출된 HTML 길이: {result.Length} 문자");
                }
                else
                {
                    Console.WriteLine("HTML 데이터 없음, 일반 텍스트 사용");
                    result = selection.Text;
                }

                // 수식이 포함된 경우 OMML을 MathML로 변환
                if (result.Contains("m:oMath"))
                {
                    Console.WriteLine("수식(OMML) 발견, MathML로 변환 시작");
                    result = ConvertOMMLToMathML(result);
                    Console.WriteLine("수식 변환 완료");
                }
                else
                {
                    Console.WriteLine("수식이 포함되지 않은 일반 텍스트");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"수식 처리 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            }
            finally
            {
                Clipboard.SetDataObject(originalClipboard);
                Console.WriteLine("클립보드 복원 완료");
            }

            return result;
        }

        private string ConvertOMMLToMathML(string ommlContent)
        {
            try
            {
                Console.WriteLine("OMML을 MathML로 변환 시작");
                
                // OMML을 MathML로 변환하는 XSLT 변환
                var xslt = new System.Xml.Xsl.XslCompiledTransform();
                var xsltPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OMML2MML.XSL");
                
                Console.WriteLine($"XSLT 파일 경로: {xsltPath}");
                
                if (File.Exists(xsltPath))
                {
                    Console.WriteLine("XSLT 파일 발견, 변환 시작");
                    xslt.Load(xsltPath);
                    
                    // OMML 문서 생성
                    var ommlDoc = new XmlDocument();
                    ommlDoc.LoadXml(ommlContent);
                    
                    // 네임스페이스 추가
                    var nsManager = new XmlNamespaceManager(ommlDoc.NameTable);
                    nsManager.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                    nsManager.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
                    
                    using (var reader = new StringReader(ommlContent))
                    using (var writer = new StringWriter())
                    {
                        var settings = new XmlReaderSettings
                        {
                            DtdProcessing = DtdProcessing.Parse
                        };
                        
                        using (var xmlReader = XmlReader.Create(reader, settings))
                        {
                            xslt.Transform(xmlReader, null, writer);
                            var result = writer.ToString();
                            Console.WriteLine($"변환 완료. 결과 길이: {result.Length} 문자");
                            return result;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("XSLT 파일을 찾을 수 없음");
                    // XSLT 파일이 없는 경우 직접 변환 시도
                    return ConvertOMMLToMathMLDirectly(ommlContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"수식 변환 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return ConvertOMMLToMathMLDirectly(ommlContent);
            }
        }

        private string ConvertOMMLToMathMLDirectly(string ommlContent)
        {
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(ommlContent);
                
                var nsManager = new XmlNamespaceManager(doc.NameTable);
                nsManager.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                
                // 분수 처리
                var fractions = doc.SelectNodes("//m:f", nsManager);
                if (fractions != null)
                {
                    foreach (XmlNode fraction in fractions)
                    {
                        var numerator = fraction.SelectSingleNode(".//m:num", nsManager)?.InnerText ?? "";
                        var denominator = fraction.SelectSingleNode(".//m:den", nsManager)?.InnerText ?? "";
                        
                        // 분수 형태로 변환
                        var mathML = $"<math xmlns='http://www.w3.org/1998/Math/MathML'><mfrac><mrow>{numerator}</mrow><mrow>{denominator}</mrow></mfrac></math>";
                        fraction.ParentNode.InnerXml = mathML;
                    }
                }
                
                return doc.InnerXml;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"직접 변환 중 오류 발생: {ex.Message}");
                return ommlContent;
            }
        }

        /// <summary>
        /// CF_HTML 포맷에서 <!--StartFragment--> ~ <!--EndFragment--> 사이만 추출
        /// </summary>
        private string ExtractHtmlFragment(string cfHtml)
        {
            if (string.IsNullOrEmpty(cfHtml))
                return string.Empty;

            // Microsoft CF_HTML spec에 따른 fragment 태그 위치 파싱
            var startMatch = Regex.Match(cfHtml, @"<!--StartFragment-->(.*)<!--EndFragment-->", RegexOptions.Singleline);
            if (startMatch.Success)
                return startMatch.Groups[1].Value.Trim();

            return cfHtml; // fallback: 전체 리턴
        }
    }
}
