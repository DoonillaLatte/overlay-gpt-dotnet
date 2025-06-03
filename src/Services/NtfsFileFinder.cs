using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;

namespace overlay_gpt.Services
{
    public class NtfsFileFinder
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int NtCreateFile(
            out IntPtr FileHandle,
            uint DesiredAccess,
            ref OBJECT_ATTRIBUTES ObjectAttributes,
            out IO_STATUS_BLOCK IoStatusBlock,
            IntPtr AllocationSize,
            uint FileAttributes,
            uint ShareAccess,
            uint CreateDisposition,
            uint CreateOptions,
            IntPtr EaBuffer,
            uint EaLength);

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int NtQueryInformationFile(
            IntPtr FileHandle,
            out IO_STATUS_BLOCK IoStatusBlock,
            IntPtr FileInformation,
            uint Length,
            int FileInformationClass);

        [StructLayout(LayoutKind.Sequential)]
        private struct OBJECT_ATTRIBUTES
        {
            public int Length;
            public IntPtr RootDirectory;
            public IntPtr ObjectName;
            public uint Attributes;
            public IntPtr SecurityDescriptor;
            public IntPtr SecurityQualityOfService;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IO_STATUS_BLOCK
        {
            public uint Status;
            public IntPtr Information;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILE_ID_BOTH_DIR_INFO
        {
            public uint NextEntryOffset;
            public uint FileIndex;
            public long CreationTime;
            public long LastAccessTime;
            public long LastWriteTime;
            public long ChangeTime;
            public long EndOfFile;
            public long AllocationSize;
            public uint FileAttributes;
            public uint FileNameLength;
            public uint EaSize;
            public byte ShortNameLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 12)]
            public byte[] ShortName;
            public long FileId;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] FileName;
        }

        private const uint FILE_OPEN = 1;
        private const uint FILE_OPEN_BY_FILE_ID = 0x00002000;
        private const uint FILE_READ_ATTRIBUTES = 0x0080;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint FILE_SHARE_DELETE = 0x00000004;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
        private const uint OBJ_CASE_INSENSITIVE = 0x00000040;
        private const int FileIdBothDirectoryInformation = 37;

        public string FindFileByFileIdAndVolumeId(long fileId, long volumeId)
        {
            try
            {
                // 볼륨 핸들 가져오기
                var volumePath = $@"\\?\Volume{{{volumeId:X8}-0000-0000-0000-000000000000}}\";
                var volumeHandle = CreateFile(
                    volumePath,
                    FILE_READ_ATTRIBUTES,
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                    IntPtr.Zero,
                    FILE_OPEN,
                    FILE_ATTRIBUTE_NORMAL,
                    IntPtr.Zero);

                if (volumeHandle.ToInt64() == -1)
                {
                    Console.WriteLine($"볼륨 핸들 열기 실패: {Marshal.GetLastWin32Error()}");
                    return null;
                }

                try
                {
                    var objectAttributes = new OBJECT_ATTRIBUTES
                    {
                        Length = Marshal.SizeOf(typeof(OBJECT_ATTRIBUTES)),
                        RootDirectory = volumeHandle,
                        ObjectName = IntPtr.Zero,
                        Attributes = OBJ_CASE_INSENSITIVE,
                        SecurityDescriptor = IntPtr.Zero,
                        SecurityQualityOfService = IntPtr.Zero
                    };

                    IntPtr fileHandle;
                    IO_STATUS_BLOCK ioStatusBlock;

                    var status = NtCreateFile(
                        out fileHandle,
                        FILE_READ_ATTRIBUTES,
                        ref objectAttributes,
                        out ioStatusBlock,
                        IntPtr.Zero,
                        0,
                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                        FILE_OPEN,
                        FILE_OPEN_BY_FILE_ID,
                        IntPtr.Zero,
                        0);

                    if (status != 0)
                    {
                        Console.WriteLine($"파일 열기 실패: {status}");
                        return null;
                    }

                    try
                    {
                        var buffer = Marshal.AllocHGlobal(1024);
                        try
                        {
                            status = NtQueryInformationFile(
                                fileHandle,
                                out ioStatusBlock,
                                buffer,
                                1024,
                                FileIdBothDirectoryInformation);

                            if (status == 0)
                            {
                                var fileInfo = (FILE_ID_BOTH_DIR_INFO)Marshal.PtrToStructure(buffer, typeof(FILE_ID_BOTH_DIR_INFO));
                                if (fileInfo.FileId == fileId)
                                {
                                    // 파일 경로 가져오기
                                    var fileName = new byte[fileInfo.FileNameLength];
                                    Marshal.Copy(buffer + Marshal.SizeOf(typeof(FILE_ID_BOTH_DIR_INFO)), fileName, 0, (int)fileInfo.FileNameLength);
                                    var path = Encoding.Unicode.GetString(fileName).TrimEnd('\0');
                                    return Path.Combine(volumePath, path);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(buffer);
                        }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(fileHandle);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(volumeHandle);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 검색 중 오류 발생: {ex.Message}");
            }

            return null;
        }

        public (ulong FileId, uint VolumeId) GetFileInfo(string filePath)
        {
            try
            {
                Console.WriteLine($"GetFileInfo 호출 - 파일 경로: {filePath}");
                
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("파일이 존재하지 않습니다.");
                    throw new FileNotFoundException("파일을 찾을 수 없습니다.", filePath);
                }

                IntPtr handle = CreateFile(
                    filePath,
                    FILE_READ_ATTRIBUTES,
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                    IntPtr.Zero,
                    FILE_OPEN,
                    FILE_ATTRIBUTE_NORMAL,
                    IntPtr.Zero);

                if (handle.ToInt64() == -1)
                {
                    Console.WriteLine($"CreateFile 실패 - 에러 코드: {Marshal.GetLastWin32Error()}");
                    throw new IOException("파일을 열 수 없습니다.");
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
                        throw new IOException("파일 정보를 가져올 수 없습니다.");
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
                throw;
            }
        }

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

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);
    }
} 