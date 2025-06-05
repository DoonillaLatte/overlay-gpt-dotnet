using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Xml;
using Microsoft.Extensions.Logging;
using System.IO.Compression;

namespace overlay_gpt
{
    [ComImport]
    [Guid("B5A7F9D1-46B7-4EFB-9F18-8F735E5A7F1A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IHwpDocument
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetText();
    }

    public class HwpContextReader : BaseContextReader
    {
        private readonly ILogger<HwpContextReader> _logger;

        public HwpContextReader(ILogger<HwpContextReader> logger)
        {
            _logger = logger;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

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

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

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

        private (ulong FileId, uint VolumeId)? GetFileId(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
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
                    return null;
                }

                try
                {
                    BY_HANDLE_FILE_INFORMATION fileInfo;
                    if (GetFileInformationByHandle(handle, out fileInfo))
                    {
                        ulong fileId = ((ulong)fileInfo.nFileIndexHigh << 32) | fileInfo.nFileIndexLow;
                        return (fileId, fileInfo.dwVolumeSerialNumber);
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
            }
            return null;
        }

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        private string ExtractHwpContent(string filePath)
        {
            try
            {
                _logger.LogInformation($"한글 파일 내용 추출 시작: {filePath}");
                
                string tempDir = Path.Combine(Path.GetTempPath(), "HwpExtract_" + Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);
                _logger.LogInformation($"임시 디렉토리 생성: {tempDir}");

                try
                {
                    string tempFile = Path.Combine(tempDir, Path.GetFileName(filePath));
                    File.Copy(filePath, tempFile, true);
                    _logger.LogInformation($"파일 복사 완료: {tempFile}");

                    string extension = Path.GetExtension(filePath).ToLower();
                    _logger.LogInformation($"파일 확장자: {extension}");

                    if (extension == ".hwpx")
                    {
                        _logger.LogInformation("HWPX 파일 처리 시작");
                        ZipFile.ExtractToDirectory(tempFile, tempDir);
                        _logger.LogInformation("ZIP 압축 해제 완료");

                        // 스타일 정보가 있는 파일들 로드
                        string contentPath = Path.Combine(tempDir, "Contents");
                        string sectionPath = Path.Combine(contentPath, "section0.xml");
                        string stylePath = Path.Combine(contentPath, "styles.xml");
                        string headerPath = Path.Combine(contentPath, "header0.xml");
                        string footerPath = Path.Combine(contentPath, "footer0.xml");

                        var xmlDoc = new XmlDocument();
                        var rootElement = xmlDoc.CreateElement("HwpDocument");
                        xmlDoc.AppendChild(rootElement);

                        // 섹션 내용 로드
                        if (File.Exists(sectionPath))
                        {
                            var sectionDoc = new XmlDocument();
                            sectionDoc.Load(sectionPath);
                            var sectionNode = xmlDoc.ImportNode(sectionDoc.DocumentElement, true);
                            rootElement.AppendChild(sectionNode);
                        }

                        // 스타일 정보 로드
                        if (File.Exists(stylePath))
                        {
                            var styleDoc = new XmlDocument();
                            styleDoc.Load(stylePath);
                            var styleNode = xmlDoc.ImportNode(styleDoc.DocumentElement, true);
                            rootElement.AppendChild(styleNode);
                        }

                        // 헤더 정보 로드
                        if (File.Exists(headerPath))
                        {
                            var headerDoc = new XmlDocument();
                            headerDoc.Load(headerPath);
                            var headerNode = xmlDoc.ImportNode(headerDoc.DocumentElement, true);
                            rootElement.AppendChild(headerNode);
                        }

                        // 푸터 정보 로드
                        if (File.Exists(footerPath))
                        {
                            var footerDoc = new XmlDocument();
                            footerDoc.Load(footerPath);
                            var footerNode = xmlDoc.ImportNode(footerDoc.DocumentElement, true);
                            rootElement.AppendChild(footerNode);
                        }

                        return xmlDoc.OuterXml;
                    }
                    else
                    {
                        _logger.LogInformation("HWP 파일 처리 시작");
                        byte[] fileContent = File.ReadAllBytes(tempFile);
                        _logger.LogInformation($"파일 크기: {fileContent.Length}바이트");
                        
                        // HWP 파일의 경우 OLE Compound File 형식이므로 기본 텍스트만 추출
                        string content = System.Text.Encoding.UTF8.GetString(fileContent);
                        var xmlDoc = new XmlDocument();
                        var rootElement = xmlDoc.CreateElement("HwpDocument");
                        var textElement = xmlDoc.CreateElement("Text");
                        textElement.InnerText = content;
                        rootElement.AppendChild(textElement);
                        xmlDoc.AppendChild(rootElement);
                        
                        return xmlDoc.OuterXml;
                    }
                }
                finally
                {
                    try
                    {
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                            _logger.LogInformation("임시 디렉토리 정리 완료");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"임시 파일 정리 중 오류: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"한글 파일 내용 추출 중 오류: {ex.Message}");
                _logger.LogError($"스택 트레이스: {ex.StackTrace}");
            }

            return string.Empty;
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                _logger.LogInformation("GetSelectedTextWithStyle 시작");
                var hwpProcesses = Process.GetProcessesByName("Hwp");
                _logger.LogInformation($"실행 중인 한글 프로세스 수: {hwpProcesses.Length}");

                if (hwpProcesses.Length == 0)
                {
                    throw new InvalidOperationException("한글(Hwp)이 실행 중이지 않습니다.");
                }

                Process? activeHwpProcess = null;
                string? filePath = null;
                foreach (var process in hwpProcesses)
                {
                    _logger.LogInformation($"프로세스 확인: ID={process.Id}, 제목={process.MainWindowTitle}");
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
                    {
                        if (process.MainWindowHandle == GetForegroundWindow())
                        {
                            _logger.LogInformation("활성 프로세스 발견");
                            activeHwpProcess = process;
                            string title = process.MainWindowTitle;
                            _logger.LogInformation($"창 제목: {title}");
                            
                            int lastDashIndex = title.LastIndexOf(" - ");
                            if (lastDashIndex > 0)
                            {
                                string fileName = title.Substring(0, lastDashIndex).Trim();
                                _logger.LogInformation($"파일명 부분: {fileName}");
                                
                                int startBracket = fileName.IndexOf('[');
                                int endBracket = fileName.IndexOf(']');
                                if (startBracket >= 0 && endBracket > startBracket)
                                {
                                    string directory = fileName.Substring(startBracket + 1, endBracket - startBracket - 1);
                                    string file = fileName.Substring(0, startBracket).Trim();
                                    filePath = Path.Combine(directory, file);
                                    _logger.LogInformation($"추출된 파일 경로: {filePath}");
                                }
                                else
                                {
                                    _logger.LogWarning("대괄호를 찾을 수 없습니다.");
                                }
                            }
                            else
                            {
                                _logger.LogWarning("대시(-)를 찾을 수 없습니다.");
                            }
                            break;
                        }
                    }
                }

                if (activeHwpProcess == null)
                {
                    _logger.LogWarning("활성 한글 프로세스를 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                if (string.IsNullOrEmpty(filePath))
                {
                    _logger.LogWarning("파일 경로를 추출할 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                try
                {
                    if (!File.Exists(filePath))
                    {
                        _logger.LogWarning($"파일을 찾을 수 없습니다: {filePath}");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    string content = ExtractHwpContent(filePath);
                    if (string.IsNullOrEmpty(content))
                    {
                        _logger.LogWarning("추출된 내용이 비어있습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    _logger.LogInformation($"추출된 내용 길이: {content.Length}자");

                    // XML 파싱
                    var xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(content);

                    // 스타일 속성 추출
                    var styleAttributes = new Dictionary<string, object>();
                    var styleNodes = xmlDoc.SelectNodes("//Style");
                    if (styleNodes != null)
                    {
                        foreach (XmlNode styleNode in styleNodes)
                        {
                            foreach (XmlAttribute attr in styleNode.Attributes)
                            {
                                styleAttributes[attr.Name] = attr.Value;
                            }
                        }
                    }

                    // 텍스트 내용 추출
                    var textBuilder = new System.Text.StringBuilder();
                    var textNodes = xmlDoc.SelectNodes("//text()");
                    if (textNodes != null)
                    {
                        foreach (XmlNode node in textNodes)
                        {
                            textBuilder.AppendLine(node.Value);
                        }
                    }

                    string result = textBuilder.ToString().Trim();
                    _logger.LogInformation($"최종 텍스트 길이: {result.Length}자");

                    return (content, styleAttributes, "1");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"한글 문서 내용 읽기 오류: {ex.Message}");
                    _logger.LogError($"스택 트레이스: {ex.StackTrace}");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"한글 데이터 읽기 오류: {ex.Message}");
                _logger.LogError($"스택 트레이스: {ex.StackTrace}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                var hwpProcesses = Process.GetProcessesByName("Hwp");
                if (hwpProcesses.Length == 0)
                {
                    return (null, null, "Hwp", string.Empty, string.Empty);
                }

                Process? activeHwpProcess = null;
                foreach (var process in hwpProcesses)
                {
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
                    {
                        if (process.MainWindowHandle == GetForegroundWindow())
                        {
                            activeHwpProcess = process;
                            break;
                        }
                    }
                }

                if (activeHwpProcess == null)
                {
                    return (null, null, "Hwp", string.Empty, string.Empty);
                }

                string filePath = activeHwpProcess.MainModule?.FileName ?? string.Empty;
                string fileName = Path.GetFileName(filePath);

                var fileIdInfo = GetFileId(filePath);

                return (
                    fileIdInfo?.FileId,
                    fileIdInfo?.VolumeId,
                    "Hwp",
                    fileName,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Hwp", string.Empty, string.Empty);
            }
        }
    }
}
