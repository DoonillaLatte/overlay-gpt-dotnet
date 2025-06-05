using System;
using System.Collections.Generic;
using System.Windows.Automation;
using Microsoft.Office.Interop.Word;
using WordFont = Microsoft.Office.Interop.Word.Font;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.IO;
using Forms = System.Windows.Forms;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Linq;
using System.Windows.Forms;
using WordApp = Microsoft.Office.Interop.Word.Application;
using HtmlDoc = HtmlAgilityPack.HtmlDocument;

namespace overlay_gpt
{
    public class WordContextWriter : IContextWriter
    {
        private WordApp? _wordApp;
        private Document? _document;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

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

        public bool OpenFile(string filePath)
        {
            try
            {
                _wordApp = (WordApp)GetActiveObject("Word.Application");
                _document = _wordApp.Documents.Open(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                if (_document == null || _wordApp == null)
                {
                    return false;
                }

                var selection = _wordApp.Selection;
                if (selection == null)
                {
                    return false;
                }

                // HTML 형식의 텍스트를 클립보드에 복사
                Clipboard.SetText(text, TextDataFormat.Html);
                
                // 클립보드의 내용을 현재 선택 영역에 붙여넣기
                selection.Paste();
                
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 적용 오류: {ex.Message}");
                return false;
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_document == null)
                {
                    return (null, null, "Word", string.Empty, string.Empty);
                }

                string filePath = _document.FullName;
                string fileName = _document.Name;
                
                if (string.IsNullOrEmpty(filePath))
                {
                    return (null, null, "Word", fileName, string.Empty);
                }
                
                var fileIdInfo = GetFileId(filePath);
                
                return (
                    fileIdInfo?.FileId,
                    fileIdInfo?.VolumeId,
                    "Word",
                    fileName,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Word", string.Empty, string.Empty);
            }
        }
    }
}
