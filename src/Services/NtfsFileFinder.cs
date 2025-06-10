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

        [StructLayout(LayoutKind.Sequential)]
        private struct UNICODE_STRING
        {
            public ushort Length;
            public ushort MaximumLength;
            public IntPtr Buffer;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILE_NAME_INFORMATION
        {
            public uint FileNameLength;
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
        private const int FileNameInformation = 9;
        private const uint OPEN_EXISTING = 3;
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;

        public string FindFileByFileIdAndVolumeId(long fileId, long volumeId)
        {
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 검색 시작 - FileId: {fileId}, VolumeId: {volumeId}");
            try
            {
                // 볼륨 ID를 올바른 형식으로 변환
                var volumeIdHex = volumeId.ToString("X8");
                var volumePath = $@"\\?\Volume{{{volumeIdHex}-0000-0000-0000-000000000000}}\";
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 볼륨 경로 생성: {volumePath}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 원본 VolumeId: {volumeId} (0x{volumeIdHex})");

                // 시스템에 마운트된 모든 볼륨 확인
                var drives = DriveInfo.GetDrives();
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 시스템에 마운트된 드라이브 목록:");
                string? foundDrivePath = null;
                foreach (var drive in drives)
                {
                    try
                    {
                        var volumeInfo = new StringBuilder(256);
                        var volumeName = new StringBuilder(256);
                        if (GetVolumeInformation(drive.RootDirectory.FullName, volumeName, volumeName.Capacity, out uint volumeSerialNumber, out uint maxComponentLength, out uint fileSystemFlags, volumeInfo, volumeInfo.Capacity))
                        {
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - 드라이브: {drive.Name}, 볼륨 ID: {volumeSerialNumber:X8}");
                            
                            // 현재 드라이브의 볼륨 ID가 일치하는지 확인
                            if (volumeSerialNumber.ToString("X8") == volumeIdHex)
                            {
                                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 일치하는 드라이브 발견: {drive.Name}");
                                foundDrivePath = drive.RootDirectory.FullName;
                                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 드라이브 경로로 변경: {foundDrivePath}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - 드라이브 {drive.Name} 정보 조회 실패: {ex.Message}");
                    }
                }

                if (foundDrivePath == null)
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 일치하는 드라이브를 찾을 수 없음");
                    return null;
                }

                // 볼륨 경로가 존재하는지 확인
                if (!Directory.Exists(foundDrivePath))
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 드라이브 경로가 존재하지 않음: {foundDrivePath}");
                    return null;
                }

                var volumeHandle = CreateFile(
                    foundDrivePath,  // 실제 찾은 드라이브 경로 사용
                    FILE_READ_ATTRIBUTES,
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    FILE_ATTRIBUTE_NORMAL | FILE_FLAG_BACKUP_SEMANTICS,
                    IntPtr.Zero);

                if (volumeHandle.ToInt64() == -1)
                {
                    var error = Marshal.GetLastWin32Error();
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 볼륨 핸들 열기 실패 - 에러 코드: {error}");
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 시도한 경로: {foundDrivePath}");
                    return null;
                }
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 볼륨 핸들 열기 성공");

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

                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 볼륨 핸들 값: {volumeHandle.ToInt64():X}");
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 ID (16진수): {fileId:X}");

                    // 파일 ID를 포함하는 UNICODE_STRING 구조체 생성
                    var fileIdString = new byte[8];
                    BitConverter.GetBytes(fileId).CopyTo(fileIdString, 0);
                    var fileIdPtr = Marshal.AllocHGlobal(fileIdString.Length);
                    Marshal.Copy(fileIdString, 0, fileIdPtr, fileIdString.Length);

                    var unicodeString = new UNICODE_STRING
                    {
                        Length = 8,
                        MaximumLength = 8,
                        Buffer = fileIdPtr
                    };

                    var unicodeStringPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(UNICODE_STRING)));
                    Marshal.StructureToPtr(unicodeString, unicodeStringPtr, false);

                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 ID 바이트 배열: {BitConverter.ToString(fileIdString)}");
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 ID 포인터 값: {fileIdPtr.ToInt64():X}");
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] UNICODE_STRING 포인터 값: {unicodeStringPtr.ToInt64():X}");

                    try
                    {
                        objectAttributes.ObjectName = unicodeStringPtr;
                        objectAttributes.Length = Marshal.SizeOf(typeof(OBJECT_ATTRIBUTES));
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] OBJECT_ATTRIBUTES 구조체 초기화 완료");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - Length: {objectAttributes.Length}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - RootDirectory: {objectAttributes.RootDirectory.ToInt64():X}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - ObjectName: {objectAttributes.ObjectName.ToInt64():X}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - Attributes: {objectAttributes.Attributes:X}");

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
                            var error = Marshal.GetLastWin32Error();
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] NtCreateFile 실패 - 상태 코드: {status}, Win32 에러: {error}");
                            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] IO_STATUS_BLOCK - Status: {ioStatusBlock.Status}, Information: {ioStatusBlock.Information.ToInt64():X}");
                            return null;
                        }

                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] NtCreateFile 성공");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 핸들 값: {fileHandle.ToInt64():X}");

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
                                    FileNameInformation);

                                if (status == 0)
                                {
                                    var fileInfo = (FILE_NAME_INFORMATION)Marshal.PtrToStructure(buffer, typeof(FILE_NAME_INFORMATION));
                                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 정보 조회 성공");
                                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - FileNameLength: {fileInfo.FileNameLength}");

                                    // 파일 경로 가져오기
                                    var fileName = new byte[fileInfo.FileNameLength];
                                    Marshal.Copy(buffer + Marshal.SizeOf(typeof(FILE_NAME_INFORMATION)), fileName, 0, (int)fileInfo.FileNameLength);
                                    var path = Encoding.Unicode.GetString(fileName, 0, (int)fileInfo.FileNameLength).TrimEnd('\0');
                                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 경로: {path}");
                                    
                                    // 경로 정규화
                                    path = path.Replace("\\", "/").TrimEnd('/');
                                    if (path.StartsWith("/"))
                                    {
                                        path = path.Substring(1);
                                    }
                                    
                                    var fullPath = Path.Combine(foundDrivePath, path);
                                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 전체 경로: {fullPath}");
                                    return fullPath;
                                }
                                else
                                {
                                    var error = Marshal.GetLastWin32Error();
                                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] NtQueryInformationFile 실패 - 상태 코드: {status}, Win32 에러: {error}");
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
                        Marshal.FreeHGlobal(fileIdPtr);
                        Marshal.FreeHGlobal(unicodeStringPtr);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 검색 중 예외 발생: {ex.Message}");
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 스택 트레이스: {ex.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 검색 중 치명적 오류 발생: {ex.Message}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 스택 트레이스: {ex.StackTrace}");
            }

            return null;
        }

        public (ulong FileId, uint VolumeId) GetFileInfo(string filePath)
        {
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] GetFileInfo 시작 - 파일 경로: {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 존재하지 않음: {filePath}");
                    throw new FileNotFoundException("파일을 찾을 수 없습니다.", filePath);
                }
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 존재 확인 완료");

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
                    var error = Marshal.GetLastWin32Error();
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] CreateFile 실패 - 에러 코드: {error}");
                    throw new IOException("파일을 열 수 없습니다.");
                }
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] CreateFile 성공");

                try
                {
                    BY_HANDLE_FILE_INFORMATION fileInfo;
                    if (GetFileInformationByHandle(handle, out fileInfo))
                    {
                        ulong fileId = ((ulong)fileInfo.nFileIndexHigh << 32) | fileInfo.nFileIndexLow;
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 정보 조회 성공:");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - FileId: {fileId}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - VolumeId: {fileInfo.dwVolumeSerialNumber}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - 파일 속성: {fileInfo.dwFileAttributes}");
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] - 파일 크기: {(fileInfo.nFileSizeHigh << 32) | fileInfo.nFileSizeLow} 바이트");
                        return (fileId, fileInfo.dwVolumeSerialNumber);
                    }
                    else
                    {
                        var error = Marshal.GetLastWin32Error();
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] GetFileInformationByHandle 실패 - 에러 코드: {error}");
                        throw new IOException("파일 정보를 가져올 수 없습니다.");
                    }
                }
                finally
                {
                    CloseHandle(handle);
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 핸들 닫기 완료");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 파일 정보 조회 중 예외 발생: {ex.Message}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 스택 트레이스: {ex.StackTrace}");
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

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool GetVolumeInformation(
            string rootPathName,
            StringBuilder volumeNameBuffer,
            int volumeNameSize,
            out uint volumeSerialNumber,
            out uint maximumComponentLength,
            out uint fileSystemFlags,
            StringBuilder fileSystemNameBuffer,
            int fileSystemNameSize);
    }
} 