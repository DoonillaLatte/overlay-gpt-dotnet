using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using System.IO;
using System.Xml;
using Microsoft.Extensions.Logging;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Reflection;
using HwpObjectLib;
using System.Threading;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net;

namespace overlay_gpt
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class HwpContextWriter : IContextWriter, IDisposable
    {
        private readonly ILogger<HwpContextWriter>? _logger;
        private string? _currentFilePath;
        private bool _isTargetProg;
        private HwpObject? _hwpApp;
        private HwpObjectWrapper? _hwpWrapper;
        private static string? _hwpDllPath;

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_CHAR = 0x0102;
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_SETTEXT = 0x000C;

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CoInitialize(IntPtr pvReserved);

        [DllImport("ole32.dll")]
        private static extern void CoUninitialize();

        [DllImport("ole32.dll")]
        private static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, int dwClsContext, ref Guid riid, out IntPtr ppv);

        [DllImport("oleaut32.dll")]
        private static extern void SysFreeString(IntPtr bstr);

        [DllImport("oleaut32.dll")]
        private static extern IntPtr SysAllocString([MarshalAs(UnmanagedType.LPWStr)] string str);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private const uint GW_CHILD = 5;
        private const uint GW_HWNDNEXT = 2;

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // IUnknown 인터페이스
        [ComImport, Guid("00000000-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IUnknown
        {
            [PreserveSig]
            int QueryInterface(ref Guid riid, out IntPtr ppvObject);
            [PreserveSig]
            int AddRef();
            [PreserveSig]
            int Release();
        }

        // IDispatch 인터페이스 GUID
        private static Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        // IDispatch 인터페이스 정의
        [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount(out int pctinfo);
            [PreserveSig]
            int GetTypeInfo(int iTInfo, int lcid, out IntPtr ppTInfo);
            [PreserveSig]
            int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames, int cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig]
            int Invoke(int dispIdMember, ref Guid riid, int lcid, short wFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, out object pVarResult, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, out int puArgErr);
        }

        // HwpObject 래퍼 클래스
        public class HwpObjectWrapper
        {
            private readonly object _comObject;
            private readonly IDispatch _dispatch;
            private readonly ILogger<HwpContextWriter>? _logger;

            public HwpObjectWrapper(object comObject, ILogger<HwpContextWriter>? logger = null)
            {
                _comObject = comObject;
                _dispatch = comObject as IDispatch;
                _logger = logger;
            }

            public string Version
            {
                get
                {
                    try
                    {
                        return InvokeMethod("Version") as string ?? "알 수 없음";
                    }
                    catch
                    {
                        return "알 수 없음";
                    }
                }
            }

            public dynamic XHwpDocuments
            {
                get
                {
                    try
                    {
                        return InvokeMethod("XHwpDocuments");
                    }
                    catch
                    {
                        return null;
                    }
                }
            }

            public void Open(string filePath, string format = "", string arg = "")
            {
                InvokeMethod("Open", filePath, format, arg);
            }

            public void RegisterModule(string module, string name)
            {
                InvokeMethod("RegisterModule", module, name);
            }

            public object CreateAction(string actionName)
            {
                return InvokeMethod("CreateAction", actionName);
            }

            // 다른 방식의 텍스트 입력 메서드들
            public void InsertText(string text)
            {
                InvokeMethod("InsertText", text);
            }

            public void PutFieldText(string text)
            {
                InvokeMethod("PutFieldText", text);
            }

            public void Run(string action)
            {
                InvokeMethod("Run", action);
            }

            public void TypeText(string text)
            {
                InvokeMethod("TypeText", text);
            }

            // 사용 가능한 메서드 목록 확인 (TypeInfo를 통한 실제 메서드 탐색)
            public string[] GetAvailableMethods()
            {
                try
                {
                    if (_dispatch == null)
                    {
                        _logger?.LogWarning("IDispatch 인터페이스가 null입니다.");
                        return new string[0];
                    }

                    var methods = new List<string>();
                    
                    // 먼저 TypeInfo를 통해 실제 사용 가능한 메서드들을 확인
                    try
                    {
                        _logger?.LogInformation("TypeInfo를 통한 실제 메서드 탐색 시도...");
                        
                        var realMethods = GetRealMethodsFromTypeInfo();
                        if (realMethods.Length > 0)
                        {
                            _logger?.LogInformation($"TypeInfo에서 실제 메서드 {realMethods.Length}개 발견: {string.Join(", ", realMethods)}");
                            methods.AddRange(realMethods);
                        }
                        else
                        {
                            _logger?.LogWarning("TypeInfo에서 메서드를 찾을 수 없습니다.");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"TypeInfo 접근 실패: {ex.Message}");
                    }
                    
                    // 한글 전용 COM 메서드들 테스트
                    string[] testMethods = { 
                        // 기본 메서드들
                        "CreateAction", "InsertText", "PutFieldText", "Run", "TypeText",
                        "SetFieldText", "AppendText", "WriteText", "Clear", "GetText",
                        "MovePos", "HAction", "DoCommand", "Execute",
                        
                        // 한글 전용 메서드들
                        "Init", "Open", "Save", "SaveAs", "Close", "Quit",
                        "XHwpDocuments", "ActiveDocument", "Documents", "Application",
                        "EditFieldName", "PutFieldText", "GetFieldText", 
                        "MoveToField", "InsertText", "SetCurPos", "GetCurPos",
                        "HParameterSet", "HAction", "HwpInsertText", "HwpEditFieldName",
                        "RegisterModule", "ReleaseModule", "HwpMessageBox",
                        "Find", "Replace", "FindReplace", "MoveNext", "MovePrev",
                        "MoveTo", "MoveToBySet", "GetSelectedPos", "SetSelectedPos",
                        "Selection", "Range", "InsertPicture", "InsertTable",
                        "CreateSet", "ReleaseSet", "HParameterAdd", "HParameterItem",
                        "HActionEx", "DoCommand2", "HwpApply", "DoMenuItem",
                        
                        // 추가 메서드들
                        "get_ActiveDocument", "get_XHwpDocuments", "get_Application",
                        "CreateParameterSet", "ModifyAction", "get_Version",
                        
                        // COM 기본 메서드들
                        "QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
                        "GetTypeInfo", "GetIDsOfNames", "Invoke",
                        
                        // 한글 COM 실제 메서드명 추정
                        "CreateInstance", "Item", "Count", "Add", "Remove"
                    };
                    
                    _logger?.LogInformation($"총 {testMethods.Length}개 메서드 존재 여부 확인 중...");
                    
                    foreach (string method in testMethods)
                    {
                        try
                        {
                            // GetIDsOfNames로 메서드 존재 여부 확인
                            string[] names = { method };
                            int[] dispIds = new int[1];
                            Guid iidNull = Guid.Empty;
                            int hr = _dispatch.GetIDsOfNames(ref iidNull, names, 1, 0, dispIds);
                            
                            if (hr == 0) // 성공하면 메서드가 존재
                            {
                                methods.Add(method);
                                _logger?.LogInformation($"메서드 발견: {method} (DispID: {dispIds[0]})");
                            }
                            else
                            {
                                _logger?.LogDebug($"메서드 없음: {method} (HRESULT: 0x{hr:X8})");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogDebug($"메서드 확인 실패: {method} - {ex.Message}");
                        }
                    }
                    
                    // 메서드가 하나도 발견되지 않으면 객체 타입 정보 출력
                    if (methods.Count == 0)
                    {
                        _logger?.LogWarning("메서드가 하나도 발견되지 않았습니다. COM 객체 정보를 확인합니다.");
                        
                        try
                        {
                            // TypeInfoCount 확인
                            int typeInfoCount;
                            int hr = _dispatch.GetTypeInfoCount(out typeInfoCount);
                            _logger?.LogInformation($"TypeInfoCount: {typeInfoCount} (HRESULT: 0x{hr:X8})");
                            
                            // 객체 타입 정보
                            _logger?.LogInformation($"COM 객체 타입: {_comObject?.GetType()?.FullName ?? "알 수 없음"}");
                            _logger?.LogInformation($"COM 객체 CLSID 추정: !HwpObject.110.1");
                            
                            // 이 객체가 실제로는 Document 객체일 수 있으므로 Application 객체 접근 시도
                            _logger?.LogInformation("이 객체가 Document 객체일 가능성 있음. Application 객체 접근 시도...");
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"객체 정보 확인 실패: {ex.Message}");
                        }
                    }
                    
                    _logger?.LogInformation($"사용 가능한 메서드 {methods.Count}개 발견: {string.Join(", ", methods)}");
                    return methods.ToArray();
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"메서드 목록 확인 중 오류: {ex.Message}");
                    return new string[0];
                }
            }

            // 한글 문서에 직접 텍스트 삽입
            public bool InsertTextToDocument(string text)
            {
                try
                {
                    _logger?.LogInformation($"텍스트 삽입 시도: {text.Substring(0, Math.Min(50, text.Length))}...");
                    
                    // 방법 0: Application 객체 접근 시도
                    try
                    {
                        _logger?.LogInformation("Application 객체 접근 시도...");
                        
                        // 현재 객체가 Document라면 Application으로 이동
                        var app = TryGetApplicationObject();
                        if (app != null)
                        {
                            _logger?.LogInformation("Application 객체 발견, Application을 통한 텍스트 삽입 시도...");
                            var appWrapper = new HwpObjectWrapper(app, _logger);
                            
                            if (appWrapper.InsertTextToDocument(text))
                            {
                                _logger?.LogInformation("Application 객체를 통한 텍스트 삽입 성공");
                                return true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"Application 객체 접근 실패: {ex.Message}");
                    }
                    
                    // 방법 1: HAction을 사용한 표준 한글 API 방식
                    try
                    {
                        _logger?.LogInformation("HAction 방식 시도 중...");
                        
                        // HAction 생성
                        var action = InvokeMethod("CreateAction", "InsertText");
                        if (action != null)
                        {
                            _logger?.LogInformation("HAction 생성 성공, 파라미터 설정 중...");
                            
                            // 파라미터 설정
                            var paramSet = InvokeMethod("CreateSet", "InsertText");
                            if (paramSet != null)
                            {
                                InvokeMethod("HParameterAdd", paramSet, "Text", text);
                                InvokeMethod("HAction", action, paramSet);
                                _logger?.LogInformation("HAction 방식으로 텍스트 삽입 성공");
                                return true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"HAction 방식 실패: {ex.Message}");
                    }

                    // 방법 2: 문서 활성화 후 직접 텍스트 삽입
                    try
                    {
                        _logger?.LogInformation("문서 활성화 후 텍스트 삽입 시도...");
                        
                        // 활성 문서 가져오기
                        var activeDoc = InvokeMethod("get_ActiveDocument");
                        if (activeDoc != null)
                        {
                            _logger?.LogInformation("활성 문서 발견, 텍스트 삽입 중...");
                            InvokeMethod("InsertText", text);
                            _logger?.LogInformation("활성 문서 방식으로 텍스트 삽입 성공");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"활성 문서 방식 실패: {ex.Message}");
                    }

                    // 방법 3: Run 메서드를 사용한 방식
                    try
                    {
                        _logger?.LogInformation("Run 메서드 방식 시도...");
                        InvokeMethod("Run", "InsertText");
                        InvokeMethod("Run", text);
                        _logger?.LogInformation("Run 방식으로 텍스트 삽입 성공");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"Run 방식 실패: {ex.Message}");
                    }

                    // 방법 4: 직접 메서드 호출
                    try
                    {
                        _logger?.LogInformation("직접 메서드 호출 시도...");
                        InvokeMethod("InsertText", text);
                        _logger?.LogInformation("직접 방식으로 텍스트 삽입 성공");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"직접 방식 실패: {ex.Message}");
                    }

                    _logger?.LogError("모든 COM 기반 텍스트 삽입 방식이 실패했습니다.");
                    return false;
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"텍스트 삽입 중 전체 오류: {ex.Message}");
                    return false;
                }
            }

            // TypeInfo에서 실제 메서드 목록 읽기
            private string[] GetRealMethodsFromTypeInfo()
            {
                try
                {
                    if (_dispatch == null)
                        return new string[0];

                    IntPtr pTypeInfo;
                    int hr = _dispatch.GetTypeInfo(0, 0, out pTypeInfo);
                    
                    if (hr != 0 || pTypeInfo == IntPtr.Zero)
                    {
                        _logger?.LogWarning($"TypeInfo 획득 실패 (HRESULT: 0x{hr:X8})");
                        return new string[0];
                    }

                    try
                    {
                        var methods = new List<string>();
                        
                        // ITypeInfo 인터페이스를 통해 메서드 정보 확인
                        // 먼저 GetIDsOfNames에서 성공하는 일반적인 메서드들을 찾아보기
                        
                        // 한글에서 실제로 사용되는 메서드명들 (한국어 포함)
                        string[] koreanMethods = {
                            "문서삽입", "텍스트삽입", "문자입력", "문자삽입",
                            "InsertText", "AddText", "WriteText", "PutText",
                            "액션생성", "액션실행", "명령실행", "실행",
                            "CreateAction", "Run", "Execute", "Do",
                            "문서", "활성문서", "현재문서",
                            "Document", "ActiveDocument", "CurrentDocument",
                            "애플리케이션", "응용프로그램",
                            "Application", "App"
                        };

                        foreach (string method in koreanMethods)
                        {
                            try
                            {
                                string[] names = { method };
                                int[] dispIds = new int[1];
                                Guid iidNull = Guid.Empty;
                                int methodHr = _dispatch.GetIDsOfNames(ref iidNull, names, 1, 0, dispIds);
                                
                                if (methodHr == 0)
                                {
                                    methods.Add(method);
                                    _logger?.LogInformation($"실제 메서드 발견: {method} (DispID: {dispIds[0]})");
                                }
                            }
                            catch
                            {
                                // 무시
                            }
                        }

                        return methods.ToArray();
                    }
                    finally
                    {
                        if (pTypeInfo != IntPtr.Zero)
                            Marshal.Release(pTypeInfo);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"TypeInfo에서 메서드 읽기 실패: {ex.Message}");
                    return new string[0];
                }
            }

            // Application 객체 접근 시도
            private object? TryGetApplicationObject()
            {
                try
                {
                    // 방법 1: Application 속성 접근
                    try
                    {
                        var app = InvokeMethod("get_Application");
                        if (app != null)
                        {
                            _logger?.LogInformation("get_Application으로 Application 객체 획득");
                            return app;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogDebug($"get_Application 실패: {ex.Message}");
                    }

                    // 방법 2: Application 직접 접근
                    try
                    {
                        var app = InvokeMethod("Application");
                        if (app != null)
                        {
                            _logger?.LogInformation("Application으로 Application 객체 획득");
                            return app;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogDebug($"Application 접근 실패: {ex.Message}");
                    }

                    // 방법 3: Parent 또는 Owner 접근
                    try
                    {
                        var parent = InvokeMethod("get_Parent");
                        if (parent != null)
                        {
                            _logger?.LogInformation("get_Parent로 상위 객체 획득");
                            return parent;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogDebug($"get_Parent 실패: {ex.Message}");
                    }

                    _logger?.LogWarning("Application 객체에 접근할 수 없습니다.");
                    return null;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Application 객체 접근 중 오류: {ex.Message}");
                    return null;
                }
            }

            public object InvokeMethod(string methodName, params object[] args)
            {
                if (_dispatch == null)
                    throw new InvalidOperationException("IDispatch 인터페이스를 사용할 수 없습니다.");

                // GetIDsOfNames로 메서드 ID 가져오기
                string[] names = { methodName };
                int[] dispIds = new int[1];
                Guid iidNull = Guid.Empty;
                int hr = _dispatch.GetIDsOfNames(ref iidNull, names, 1, 0, dispIds);
                
                if (hr != 0)
                    throw new Exception($"메서드 '{methodName}'를 찾을 수 없습니다. (HRESULT: 0x{hr:X8})");

                // DISPPARAMS 구조체 설정
                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS();
                dispParams.cArgs = args.Length;
                
                if (args.Length > 0)
                {
                    // 인수들을 VARIANT 배열로 변환
                    dispParams.rgvarg = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(object)) * args.Length);
                    for (int i = 0; i < args.Length; i++)
                    {
                        Marshal.GetNativeVariantForObject(args[args.Length - 1 - i], 
                            IntPtr.Add(dispParams.rgvarg, i * Marshal.SizeOf(typeof(object))));
                    }
                }

                // Invoke 호출
                object result;
                var excepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                int argErr;
                
                hr = _dispatch.Invoke(dispIds[0], ref iidNull, 0, 1, ref dispParams, out result, ref excepInfo, out argErr);
                
                // 메모리 해제
                if (dispParams.rgvarg != IntPtr.Zero)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        Marshal.FreeCoTaskMem(IntPtr.Add(dispParams.rgvarg, i * Marshal.SizeOf(typeof(object))));
                    }
                    Marshal.FreeCoTaskMem(dispParams.rgvarg);
                }

                if (hr != 0)
                    throw new Exception($"메서드 '{methodName}' 호출 실패. (HRESULT: 0x{hr:X8})");

                return result;
            }
        }

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        public bool IsTargetProg
        {
            get => _isTargetProg;
            set => _isTargetProg = value;
        }

        public HwpContextWriter(ILogger<HwpContextWriter>? logger = null, bool isTargetProg = false)
        {
            if (logger == null)
            {
                var factory = LoggerFactory.Create(builder => builder.AddConsole());
                _logger = factory.CreateLogger<HwpContextWriter>();
            }
            else
            {
                _logger = logger;
            }
            _isTargetProg = isTargetProg;
        }

        private HwpObject? FindHwpInROT()
        {
            try
            {
                IRunningObjectTable rot;
                IEnumMoniker enumMoniker;
                int result = GetRunningObjectTable(0, out rot);
                
                if (result != 0)
                {
                    _logger?.LogError($"ROT 접근 실패 (코드: {result}). 관리자 권한이 필요할 수 있습니다.");
                    return null;
                }

                rot.EnumRunning(out enumMoniker);

                IMoniker[] monikers = new IMoniker[1];
                IntPtr fetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, fetched) == 0)
                {
                    try
                    {
                        IBindCtx bindCtx;
                        CreateBindCtx(0, out bindCtx);
                        string displayName;
                        monikers[0].GetDisplayName(bindCtx, null, out displayName);

                        _logger?.LogInformation($"ROT에서 찾은 객체: {displayName}");

                        // 한글 객체 검색 조건 확장
                        if (displayName.Contains("!HwpObject") || 
                            displayName.Contains("HWPFrame.HwpObject") ||
                            displayName.Contains("!HwpApplication") ||
                            displayName.Contains("!Hwp.Application") ||
                            (displayName.Contains("Hwp") && displayName.Contains("!")))
                        {
                            try
                            {
                                _logger?.LogInformation($"한글 객체 후보 발견, 연결 시도: {displayName}");
                                object obj;
                                rot.GetObject(monikers[0], out obj);
                                
                                _logger?.LogInformation($"ROT 객체 타입: {obj?.GetType()?.FullName ?? "null"}");
                                
                                if (obj is HwpObject hwpObj)
                                {
                                    try
                                    {
                                        // 객체가 유효한지 테스트
                                        var version = hwpObj.Version;
                                        _logger?.LogInformation($"ROT에서 한글 객체를 찾았습니다. (버전: {version})");
                                        return hwpObj;
                                    }
                                    catch
                                    {
                                        _logger?.LogWarning("ROT에서 찾은 객체가 유효하지 않습니다.");
                                        Marshal.ReleaseComObject(obj);
                                    }
                                }
                                else
                                {
                                    // IDispatch 래퍼를 통한 접근 시도
                                    try
                                    {
                                        _logger?.LogInformation("IDispatch 래퍼를 통한 COM 객체 접근 시도...");
                                        
                                        var wrapper = new HwpObjectWrapper(obj, _logger);
                                        string version = wrapper.Version;
                                        _logger?.LogInformation($"IDispatch 래퍼 접근 성공, 버전: {version}");
                                        
                                        // 래퍼 객체 저장
                                        _hwpWrapper = wrapper;
                                        _logger?.LogInformation("ROT에서 한글 객체를 IDispatch 래퍼로 연결했습니다.");
                                        
                                        // 래퍼가 성공적으로 생성되었으므로 dummy HwpObject 반환
                                        // 실제 작업은 _hwpWrapper를 통해 수행됨
                                        return new object() as HwpObject ?? obj as HwpObject;
                                    }
                                    catch (Exception wrapperEx)
                                    {
                                        _logger?.LogWarning($"IDispatch 래퍼 접근 실패: {wrapperEx.Message}");
                                        
                                        // 기존 dynamic 접근 방식으로 폴백
                                        try
                                        {
                                            _logger?.LogInformation("Dynamic 접근으로 폴백...");
                                            dynamic dynObj = obj;
                                            string version = dynObj.Version;
                                            _logger?.LogInformation($"Dynamic 접근 성공, 버전: {version}");
                                            return obj as HwpObject;
                                        }
                                        catch (Exception dynEx)
                                        {
                                            _logger?.LogWarning($"Dynamic 접근도 실패: {dynEx.Message}");
                                            
                                            // 최종적으로 COM 객체 해제
                                            try
                                            {
                                                Marshal.ReleaseComObject(obj);
                                            }
                                            catch
                                            {
                                                // COM 객체 해제 실패는 무시
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"ROT 객체 접근 실패: {ex.Message}");
                            }
                        }

                        // 리소스 해제
                        Marshal.ReleaseComObject(bindCtx);
                        Marshal.ReleaseComObject(monikers[0]);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"ROT 항목 처리 중 오류: {ex.Message}");
                    }
                }

                // 리소스 해제
                Marshal.ReleaseComObject(enumMoniker);
                Marshal.ReleaseComObject(rot);

                _logger?.LogInformation("ROT에서 한글 객체를 찾지 못했습니다.");
                return null;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"ROT 검색 중 오류 발생: {ex.Message}");
                return null;
            }
        }

        private bool InitializeHwpApplication()
        {
            try
            {
                _logger?.LogInformation("한글 초기화 시도...");
                
                // COM 라이브러리 초기화
                int comResult = CoInitialize(IntPtr.Zero);
                _logger?.LogInformation($"COM 초기화 결과: {comResult}");

                // 1. 먼저 실행 중인 한글 프로세스 찾기
                var processes = Process.GetProcessesByName("Hwp");
                if (processes.Length > 0)
                {
                    _logger?.LogInformation($"실행 중인 한글 프로세스 발견: {processes[0].ProcessName} (PID: {processes[0].Id})");
                    _logger?.LogInformation($"프로세스 경로: {processes[0].MainModule?.FileName ?? "알 수 없음"}");
                    
                    // ROT에서 먼저 찾기 시도
                    _hwpApp = FindHwpInROT();
                    if (_hwpApp != null || _hwpWrapper != null)
                    {
                        _logger?.LogInformation($"ROT 연결 성공 - HwpApp: {_hwpApp != null}, Wrapper: {_hwpWrapper != null}");
                        return true;
                    }

                    // COM 객체 연결 시도 (최대 10회)
                    for (int retryCount = 0; retryCount < 10; retryCount++)
                    {
                        Thread.Sleep(2000); // 2초씩 대기
                        _logger?.LogInformation($"실행 중인 한글 COM 객체 연결 시도 {retryCount + 1}/10...");

                        try
                        {
                            // ProgID 확인
                            _logger?.LogInformation("ProgID 'HWPFrame.HwpObject' 확인 중...");
                            Type? hwpType = Type.GetTypeFromProgID("HWPFrame.HwpObject");
                            if (hwpType != null)
                            {
                                _logger?.LogInformation($"ProgID 타입 발견: {hwpType.FullName}");
                                
                                try
                                {
                                    _logger?.LogInformation("GetActiveObject 시도 중...");
                                    Guid clsid;
                                    CLSIDFromProgID("HWPFrame.HwpObject", out clsid);
                                    _logger?.LogInformation($"CLSID: {clsid}");
                                    
                                    object obj;
                                    int result = GetActiveObject(ref clsid, IntPtr.Zero, out obj);
                                    _logger?.LogInformation($"GetActiveObject 결과: {result}");
                                    
                                    if (result >= 0)
                                    {
                                        _hwpApp = obj as HwpObject;
                                        if (_hwpApp != null)
                                        {
                                            string version = _hwpApp.Version;
                                            _logger?.LogInformation($"GetActiveObject로 COM 객체 연결 성공 (버전: {version})");
                                            return true;
                                        }
                                        else
                                        {
                                            _logger?.LogWarning($"GetActiveObject는 성공했지만 HwpObject로 형변환 실패. 실제 타입: {obj?.GetType()?.FullName ?? "null"}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogInformation($"GetActiveObject 실패: {ex.Message}, CreateInstance 시도...");
                                }

                                // CreateInstance로 시도
                                try
                                {
                                    _logger?.LogInformation("CreateInstance 시도 중...");
                                    object? raw = Activator.CreateInstance(hwpType);
                                    _logger?.LogInformation($"CreateInstance 결과 타입: {raw?.GetType()?.FullName ?? "null"}");
                                    
                                    _hwpApp = raw as HwpObject;

                                    if (_hwpApp == null)
                                    {
                                        _logger?.LogWarning($"CreateInstance로 HwpObject를 생성했으나 형변환 실패 (실제 타입: {raw?.GetType()?.FullName ?? "null"})");
                                        continue;
                                    }

                                    string hwpVersion = _hwpApp.Version;
                                    _logger?.LogInformation($"CreateInstance로 COM 객체 연결 성공 (버전: {hwpVersion})");
                                    return true;
                                }
                                catch (Exception ex2)
                                {
                                    _logger?.LogWarning($"CreateInstance 실패: {ex2.Message}");
                                }
                            }
                            else
                            {
                                _logger?.LogWarning("ProgID 'HWPFrame.HwpObject'를 찾을 수 없습니다. 다른 ProgID를 시도합니다.");
                                
                                // 레지스트리에서 찾은 ProgID들 시도
                                string[] progIds = FindHwpProgIDs();
                                _logger?.LogInformation($"시도할 ProgID 개수: {progIds.Length}");
                                foreach (string progId in progIds)
                                {
                                    _logger?.LogInformation($"ProgID '{progId}' 시도 중...");
                                    Type? alternativeType = Type.GetTypeFromProgID(progId);
                                    if (alternativeType != null)
                                    {
                                        _logger?.LogInformation($"대체 ProgID 발견: {progId} -> {alternativeType.FullName}");
                                        try
                                        {
                                            object? altObj = Activator.CreateInstance(alternativeType);
                                            _hwpApp = altObj as HwpObject;
                                            if (_hwpApp != null)
                                            {
                                                string altVersion = _hwpApp.Version;
                                                _logger?.LogInformation($"대체 ProgID로 COM 객체 연결 성공 (버전: {altVersion})");
                                                return true;
                                            }
                                        }
                                        catch (Exception ex3)
                                        {
                                            _logger?.LogWarning($"대체 ProgID '{progId}' 실패: {ex3.Message}");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"COM 객체 연결 시도 {retryCount + 1} 실패: {ex.Message}");
                            continue;
                        }
                    }
                    
                    _logger?.LogError("실행 중인 한글에 연결 실패");
                    
                    // COM 라이브러리 수동 등록 시도
                    _logger?.LogInformation("COM 라이브러리 수동 등록을 시도합니다...");
                    if (TryRegisterHwpCom())
                    {
                        _logger?.LogInformation("COM 등록 후 재시도...");
                        Thread.Sleep(3000); // 등록 후 잠시 대기
                        
                        // 등록 후 한 번 더 시도
                        for (int retryCount = 0; retryCount < 3; retryCount++)
                        {
                            Thread.Sleep(2000);
                            _logger?.LogInformation($"COM 등록 후 연결 시도 {retryCount + 1}/3...");
                            
                            try
                            {
                                Type? hwpType = Type.GetTypeFromProgID("HWPFrame.HwpObject");
                                if (hwpType != null)
                                {
                                    object? raw = Activator.CreateInstance(hwpType);
                                    _hwpApp = raw as HwpObject;
                                    if (_hwpApp != null)
                                    {
                                        string version = _hwpApp.Version;
                                        _logger?.LogInformation($"COM 등록 후 연결 성공 (버전: {version})");
                                        return true;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"COM 등록 후 연결 시도 {retryCount + 1} 실패: {ex.Message}");
                            }
                        }
                    }
                    
                    return false;
                }

                // 2. 실행 중인 프로세스가 없는 경우에만 새로 실행
                _logger?.LogInformation("실행 중인 한글 프로세스가 없습니다. 한글 프로그램을 찾아서 실행합니다.");
                
                // 다양한 한글 설치 경로 확인
                string[] hwpPaths = {
                    @"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp.exe",
                    @"C:\Program Files\Hnc\Office 2020\HOffice110\Bin\Hwp.exe",
                    @"C:\Program Files (x86)\Hnc\Office 2022\HOffice120\Bin\Hwp.exe",
                    @"C:\Program Files\Hnc\Office 2022\HOffice120\Bin\Hwp.exe",
                    @"C:\Program Files (x86)\Hnc\HwpOffice\Bin\Hwp.exe",
                    @"C:\Program Files\Hnc\HwpOffice\Bin\Hwp.exe"
                };

                string? hwpPath = null;
                foreach (string path in hwpPaths)
                {
                    _logger?.LogInformation($"한글 경로 확인: {path}");
                    if (File.Exists(path))
                    {
                        hwpPath = path;
                        _logger?.LogInformation($"한글 프로그램 발견: {hwpPath}");
                        break;
                    }
                }

                if (hwpPath == null)
                {
                    _logger?.LogError("한글 프로그램을 찾을 수 없습니다. 설치된 경로를 확인해주세요.");
                    return false;
                }

                try
                {
                    _logger?.LogInformation($"한글 프로그램 실행: {hwpPath}");
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = hwpPath,
                        UseShellExecute = true,
                        Verb = "runas"  // 관리자 권한으로 실행
                    });
                    _logger?.LogInformation("관리자 권한으로 한글 프로그램을 실행했습니다. COM 객체 초기화를 기다립니다...");

                    // COM 객체 연결 시도 (최대 10회)
                    for (int retryCount = 0; retryCount < 10; retryCount++)
                    {
                        Thread.Sleep(1000); // 1초씩 대기

                        processes = Process.GetProcessesByName("Hwp");
                        if (processes.Length == 0)
                        {
                            _logger?.LogError("한글 프로그램이 실행되지 않았습니다.");
                            return false;
                        }

                        _logger?.LogInformation($"COM 객체 연결 시도 {retryCount + 1}/10...");

                        // ROT에서 먼저 찾기 시도
                        _hwpApp = FindHwpInROT();
                        if (_hwpApp != null)
                        {
                            return true;
                        }

                        try
                        {
                            Type? hwpType = Type.GetTypeFromProgID("HWPFrame.HwpObject");
                            if (hwpType != null)
                            {
                                _hwpApp = (HwpObject)Activator.CreateInstance(hwpType);
                                if (_hwpApp != null)
                                {
                                    var version = _hwpApp.Version;
                                    _logger?.LogInformation($"COM 객체 연결 성공 (버전: {version})");
                                    return true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }

                    _logger?.LogError("최대 시도 횟수를 초과했습니다. COM 객체 연결 실패.");
                    return false;
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"한글 COM 객체 초기화 실패: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"한글 초기화 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                _logger?.LogInformation($"파일 열기 시도: {filePath}");

                // 파일 경로 정규화
                string normalizedPath = Path.GetFullPath(filePath);
                _logger?.LogInformation($"정규화된 경로: {normalizedPath}");

                // 실행 파일인 경우 새 문서 생성
                if (normalizedPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogInformation($"실행 파일이므로 새 문서를 생성합니다: {normalizedPath}");
                    // 한글 초기화만 하고 새 문서는 생성하지 않음
                    if (!InitializeHwpApplication() || (_hwpApp == null && _hwpWrapper == null))
                    {
                        _logger?.LogError("한글 초기화 실패");
                        return false;
                    }
                    _currentFilePath = "새 문서";
                    return true;
                }

                if (!File.Exists(normalizedPath))
                {
                    _logger?.LogError($"파일이 존재하지 않습니다: {normalizedPath}");
                    return false;
                }

                // 한글 초기화
                if (!InitializeHwpApplication() || (_hwpApp == null && _hwpWrapper == null))
                {
                    _logger?.LogError("한글 초기화 실패");
                    return false;
                }

                try
                {
                    _logger?.LogInformation("파일 열기 시작...");
                    
                    // 래퍼 또는 직접 객체 사용
                    if (_hwpWrapper != null)
                    {
                        _logger?.LogInformation("IDispatch 래퍼를 사용하여 파일 열기...");
                        
                        try
                        {
                            // 이미 열려있는 문서 확인 (래퍼 사용)
                            var documents = _hwpWrapper.XHwpDocuments;
                            if (documents != null)
                            {
                                _logger?.LogInformation("문서 목록 확인 중...");
                                // 래퍼를 통한 문서 확인은 복잡하므로 일단 직접 열기 시도
                            }
                            
                            // 파일 열기 전에 기본 모듈 등록
                            _hwpWrapper.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");

                            // 공백 포함 경로 대비해 따옴표로 감쌈
                            string quotedPath = $"\"{normalizedPath}\"";

                            _hwpWrapper.Open(quotedPath, "", "forceopen:true");
                            _currentFilePath = normalizedPath;
                            _logger?.LogInformation("한글 파일 열기 성공 (래퍼 사용)");
                            return true;
                        }
                        catch (Exception wrapperEx)
                        {
                            _logger?.LogError($"래퍼를 통한 파일 열기 실패: {wrapperEx.Message}");
                            return false;
                        }
                    }
                    else if (_hwpApp != null)
                    {
                        _logger?.LogInformation("직접 HwpObject를 사용하여 파일 열기...");
                        
                        // 이미 열려있는 문서 확인
                        for (int i = 0; i < _hwpApp.XHwpDocuments.Count; i++)
                        {
                            var doc = _hwpApp.XHwpDocuments.Item(i);
                            string docPath = Path.Combine(doc.Path, doc.Name);
                            _logger?.LogInformation($"열린 문서 확인: {docPath}");
                            
                            if (docPath.Equals(normalizedPath, StringComparison.OrdinalIgnoreCase))
                            {
                                _currentFilePath = normalizedPath;
                                doc.SetActive();
                                _logger?.LogInformation("이미 열려있는 한글 파일을 사용합니다.");
                                return true;
                            }
                        }

                        // 새로 파일 열기
                        _logger?.LogInformation("새로 파일 열기 시도...");
                        
                        // 파일 열기 전에 기본 모듈 등록
                        _hwpApp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");

                        // 공백 포함 경로 대비해 따옴표로 감쌈
                        string quotedPath = $"\"{normalizedPath}\"";

                        try
                        {
                            _hwpApp.Open(quotedPath, "", "forceopen:true");  // 첫 번째 시도
                            _currentFilePath = normalizedPath;
                            _logger?.LogInformation("한글 파일 열기 성공");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogError($"파일 열기 실패 (첫 번째 시도): {ex.Message}");
                            try
                            {
                                _hwpApp.Open(quotedPath, "HWP", "");  // 두 번째 시도
                                _currentFilePath = normalizedPath;
                                _logger?.LogInformation("한글 파일 열기 성공 (두 번째 시도)");
                                return true;
                            }
                            catch (Exception ex2)
                            {
                                _logger?.LogError($"파일 열기 실패 (두 번째 시도): {ex2.Message}");
                                return false;
                            }
                        }
                    }
                    else
                    {
                        _logger?.LogError("한글 객체와 래퍼가 모두 null입니다.");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"한글 파일 작업 중 오류 발생: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"한글 파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                _logger?.LogInformation($"스타일 적용된 텍스트 삽입 시도: {text.Substring(0, Math.Min(100, text.Length))}...");

                // 1. HTML/XML 형식인지 확인하고 적절히 변환
                string processedText = ProcessFormattedText(text);
                _logger?.LogInformation($"처리된 텍스트 길이: {processedText.Length}");

                if (_hwpWrapper != null)
                {
                    try
                    {
                        // 먼저 사용 가능한 메서드 확인
                        _logger?.LogInformation("사용 가능한 메서드 확인 중...");
                        var availableMethods = _hwpWrapper.GetAvailableMethods();
                        _logger?.LogInformation($"사용 가능한 메서드들: {string.Join(", ", availableMethods)}");

                        // 1. HTML 클립보드 방식으로 서식 유지하며 삽입
                        _logger?.LogInformation("HTML 클립보드 방식으로 서식 적용 시도...");
                        try
                        {
                            if (InsertFormattedTextViaClipboard(processedText))
                            {
                                _logger?.LogInformation($"서식 적용된 텍스트 삽입 완료 (HTML 클립보드): {lineNumber}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"HTML 클립보드 방식 실패: {ex.Message}");
                        }

                        // 2. 한글 HAction을 통한 서식 적용
                        _logger?.LogInformation("한글 HAction을 통한 서식 적용 시도...");
                        try
                        {
                            if (InsertFormattedTextViaHAction(processedText))
                            {
                                _logger?.LogInformation($"서식 적용된 텍스트 삽입 완료 (HAction): {lineNumber}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"HAction 서식 적용 실패: {ex.Message}");
                        }

                        // 3. 한글 문서에 직접 텍스트 삽입 시도 (폴백)
                        _logger?.LogInformation("한글 문서 직접 조작 방식 시도...");
                        try
                        {
                            if (_hwpWrapper.InsertTextToDocument(processedText))
                            {
                                _logger?.LogInformation($"텍스트 적용 완료 (문서 직접 조작): {lineNumber}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"문서 직접 조작 실패: {ex.Message}");
                        }

                        // 2. HAction 방식 시도 (한글 표준 방식)
                        if (availableMethods.Contains("HAction"))
                        {
                            try
                            {
                                _logger?.LogInformation("HAction 방식 시도...");
                                _hwpWrapper.InvokeMethod("HAction", "InsertText");
                                _logger?.LogInformation($"텍스트 적용 완료 (HAction): {lineNumber}");
                                return true;
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"HAction 방식 실패: {ex.Message}");
                            }
                        }

                        // 3. CreateAction + Run 방식 시도
                        if (availableMethods.Contains("CreateAction"))
                        {
                            try
                            {
                                _logger?.LogInformation("CreateAction 방식 시도...");
                                dynamic action = _hwpWrapper.CreateAction("InsertText");
                                action.Run(text);
                                _logger?.LogInformation($"텍스트 적용 완료 (CreateAction): {lineNumber}");
                                return true;
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"CreateAction 방식 실패: {ex.Message}");
                            }
                        }

                        // 4. 직접 InsertText 시도
                        if (availableMethods.Contains("InsertText"))
                        {
                            try
                            {
                                _logger?.LogInformation("직접 InsertText 시도...");
                                _hwpWrapper.InsertText(text);
                                _logger?.LogInformation($"텍스트 적용 완료 (InsertText): {lineNumber}");
                                return true;
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"InsertText 방식 실패: {ex.Message}");
                            }
                        }

                        // 5. PutFieldText 시도
                        if (availableMethods.Contains("PutFieldText"))
                        {
                            try
                            {
                                _logger?.LogInformation("PutFieldText 시도...");
                                _hwpWrapper.PutFieldText(text);
                                _logger?.LogInformation($"텍스트 적용 완료 (PutFieldText): {lineNumber}");
                                return true;
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"PutFieldText 방식 실패: {ex.Message}");
                            }
                        }

                        // 6. 직접 문서 조작 방식 시도 (포커스 독립적)
                        _logger?.LogInformation("직접 문서 조작 방식 시도...");
                        try
                        {
                            if (TryDirectDocumentInsertion(text))
                            {
                                _logger?.LogInformation($"텍스트 적용 완료 (직접 조작): {lineNumber}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"직접 문서 조작 방식 실패: {ex.Message}");
                        }

                        // 7. 새로운 Application 인스턴스 생성 시도 (최후의 방법)
                        _logger?.LogInformation("새로운 Application 인스턴스 생성 시도...");
                        try
                        {
                            if (TryCreateNewApplicationInstance(text))
                            {
                                _logger?.LogInformation($"텍스트 적용 완료 (새 Application): {lineNumber}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"새 Application 인스턴스 방식 실패: {ex.Message}");
                        }

                        _logger?.LogError("모든 텍스트 입력 방식이 실패했습니다.");
                        return false;
                    }
                    catch (Exception wrapperEx)
                    {
                        _logger?.LogError($"래퍼를 통한 텍스트 적용 실패: {wrapperEx.Message}");
                        return false;
                    }
                }
                else if (_hwpApp != null)
                {
                    // 직접 객체를 통한 텍스트 적용
                    dynamic action = _hwpApp.CreateAction("InsertText");
                    action.Run(text);
                    _logger?.LogInformation($"텍스트 적용 완료: {lineNumber}");
                    return true;
                }
                else
                {
                    _logger?.LogError("한글 애플리케이션이 초기화되지 않았습니다.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"텍스트 적용 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                _logger?.LogWarning("현재 파일 경로가 설정되지 않았습니다.");
                return (null, null, "Hwp", "", "");
            }

            return (null, null, "Hwp", Path.GetFileName(_currentFilePath), _currentFilePath);
        }

        private bool SendTextToHwpWindow(string text)
        {
            try
            {
                _logger?.LogInformation("한글 프로세스 및 편집 윈도우 찾기...");
                
                var processes = Process.GetProcessesByName("Hwp");
                if (processes.Length == 0)
                {
                    _logger?.LogWarning("실행 중인 한글 프로세스를 찾을 수 없습니다.");
                    return false;
                }

                IntPtr mainWindow = processes[0].MainWindowHandle;
                if (mainWindow == IntPtr.Zero)
                {
                    _logger?.LogWarning("한글 메인 윈도우를 찾을 수 없습니다.");
                    return false;
                }

                _logger?.LogInformation($"한글 메인 윈도우 찾음: {mainWindow}");

                // 1. 먼저 메인 윈도우에 포커스 설정
                SetForegroundWindow(mainWindow);
                Thread.Sleep(200);

                // 2. 여러 방법으로 텍스트 전송 시도
                
                // 방법 1: SendKeys를 통한 클립보드 붙여넣기
                try
                {
                    _logger?.LogInformation("SendKeys를 통한 클립보드 붙여넣기 시도...");
                    
                    // 클립보드에 텍스트 복사
                    Clipboard.SetText(text);
                    Thread.Sleep(200);
                    
                    // SendKeys로 Ctrl+V 전송
                    SendKeys.SendWait("^v");
                    Thread.Sleep(200);
                    
                    _logger?.LogInformation("SendKeys 클립보드 방식 완료");
                    return true;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"SendKeys 클립보드 방식 실패: {ex.Message}");
                }

                // 방법 2: 직접 타이핑 시뮬레이션
                try
                {
                    _logger?.LogInformation("SendKeys 직접 타이핑 시도...");
                    
                    // 특수 문자 이스케이프 처리
                    string escapedText = text.Replace("^", "{{^}}")
                                           .Replace("%", "{{%}}")
                                           .Replace("~", "{{~}}")
                                           .Replace("+", "{{+}}")
                                           .Replace("{", "{{{}}")
                                           .Replace("}", "{{}}}")
                                           .Replace("(", "{(}")
                                           .Replace(")", "{)}");
                    
                    SendKeys.SendWait(escapedText);
                    Thread.Sleep(100);
                    
                    _logger?.LogInformation("SendKeys 직접 타이핑 완료");
                    return true;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"SendKeys 직접 타이핑 실패: {ex.Message}");
                }

                // 방법 3: 편집 가능한 자식 윈도우 찾기
                try
                {
                    _logger?.LogInformation("편집 윈도우 찾기 시도...");
                    IntPtr editWindow = FindEditWindow(mainWindow);
                    if (editWindow != IntPtr.Zero)
                    {
                        _logger?.LogInformation($"편집 윈도우 찾음: {editWindow}");
                        SetForegroundWindow(editWindow);
                        Thread.Sleep(100);

                        // WM_SETTEXT로 직접 텍스트 설정
                        SendMessage(editWindow, WM_SETTEXT, IntPtr.Zero, text);
                        Thread.Sleep(100);
                        
                        _logger?.LogInformation("편집 윈도우 방식 완료");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"편집 윈도우 방식 실패: {ex.Message}");
                }

                // 방법 4: 문자 하나씩 전송 (개선된 방식)
                try
                {
                    _logger?.LogInformation("개선된 문자별 전송 시도...");
                    
                    foreach (char c in text)
                    {
                        // 한글 문자 처리
                        if (c >= 0xAC00 && c <= 0xD7AF) // 한글 완성형 범위
                        {
                            // 한글 문자는 UTF-16으로 처리
                            byte[] bytes = Encoding.Unicode.GetBytes(new char[] { c });
                            ushort charCode = BitConverter.ToUInt16(bytes, 0);
                            PostMessage(mainWindow, WM_CHAR, (IntPtr)charCode, IntPtr.Zero);
                        }
                        else
                        {
                            // 영어/숫자는 ASCII 코드로 처리
                            PostMessage(mainWindow, WM_CHAR, (IntPtr)c, IntPtr.Zero);
                        }
                        Thread.Sleep(50); // 문자 사이 지연 증가
                    }
                    
                    _logger?.LogInformation("문자별 전송 완료");
                    return true;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"문자별 전송 실패: {ex.Message}");
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Windows API 텍스트 전송 중 오류: {ex.Message}");
                return false;
            }
        }

        private IntPtr FindEditWindow(IntPtr parentWindow)
        {
            IntPtr editWindow = IntPtr.Zero;
            
            EnumChildWindows(parentWindow, (hWnd, lParam) =>
            {
                var className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);
                string classNameStr = className.ToString();
                
                _logger?.LogInformation($"자식 윈도우 클래스: {classNameStr}");
                
                // 편집 가능한 윈도우 클래스들
                if (classNameStr.Contains("Edit") || 
                    classNameStr.Contains("RichEdit") || 
                    classNameStr.Contains("RICHEDIT") ||
                    classNameStr.Contains("Scintilla") ||
                    classNameStr.Contains("HwpEdit"))
                {
                    editWindow = hWnd;
                    return false; // 찾았으므로 열거 중단
                }
                
                return true; // 계속 찾기
            }, IntPtr.Zero);
            
            return editWindow;
        }

        // 한글 문서에 직접 텍스트 삽입 (포커스 독립적)
        private bool TryDirectDocumentInsertion(string text)
        {
            try
            {
                _logger?.LogInformation("한글 문서 직접 조작 방식 시도...");
                
                var processes = Process.GetProcessesByName("Hwp");
                if (processes.Length == 0)
                {
                    _logger?.LogWarning("실행 중인 한글 프로세스를 찾을 수 없습니다.");
                    return false;
                }

                IntPtr mainWindow = processes[0].MainWindowHandle;
                if (mainWindow == IntPtr.Zero)
                {
                    _logger?.LogWarning("한글 메인 윈도우를 찾을 수 없습니다.");
                    return false;
                }

                _logger?.LogInformation($"한글 메인 윈도우 발견: {mainWindow}");

                // 1. 한글의 문서 편집 영역 찾기
                IntPtr docWindow = FindHwpDocumentWindow(mainWindow);
                if (docWindow != IntPtr.Zero)
                {
                    _logger?.LogInformation($"한글 문서 편집 영역 발견: {docWindow}");
                    
                                         // 2. 다양한 방법으로 텍스트 삽입 시도
                     _logger?.LogInformation("다양한 방법으로 텍스트 삽입 시도...");
                     
                     // 방법 2-1: SendKeys를 사용한 안정적인 방식
                     try
                     {
                         _logger?.LogInformation("SendKeys를 통한 텍스트 삽입 시도...");
                         
                         // 문서 영역에 포커스 설정
                         SetForegroundWindow(docWindow);
                         Thread.Sleep(300);
                         
                         // 클립보드에 텍스트 설정
                         Clipboard.SetText(text);
                         Thread.Sleep(200);
                         
                         // SendKeys로 Ctrl+V 전송 (더 안정적)
                         SendKeys.SendWait("^v");
                         Thread.Sleep(200);
                         
                         _logger?.LogInformation("SendKeys 방식 완료");
                         return true;
                     }
                     catch (Exception ex)
                     {
                         _logger?.LogWarning($"SendKeys 방식 실패: {ex.Message}");
                     }
                     
                     // 방법 2-2: 마우스 클릭으로 확실한 포커스 설정 후 삽입
                     try
                     {
                         _logger?.LogInformation("마우스 클릭 + 텍스트 삽입 시도...");
                         
                         // 문서 영역 클릭하여 확실한 포커스 설정
                         var rect = GetWindowRect(docWindow);
                         if (rect.HasValue)
                         {
                             int centerX = (rect.Value.Left + rect.Value.Right) / 2;
                             int centerY = (rect.Value.Top + rect.Value.Bottom) / 2;
                             
                             _logger?.LogInformation($"문서 영역 클릭: ({centerX}, {centerY})");
                             
                             // 마우스 클릭
                             ClickWindow(centerX, centerY);
                             Thread.Sleep(200);
                             
                             // 클립보드 붙여넣기
                             Clipboard.SetText(text);
                             Thread.Sleep(100);
                             SendKeys.SendWait("^v");
                             Thread.Sleep(200);
                             
                             _logger?.LogInformation("마우스 클릭 방식 완료");
                             return true;
                         }
                     }
                     catch (Exception ex)
                     {
                         _logger?.LogWarning($"마우스 클릭 방식 실패: {ex.Message}");
                     }
                     
                     // 방법 2-3: 직접 키보드 메시지 (개선된 방식)
                     try
                     {
                         _logger?.LogInformation("개선된 키보드 메시지 방식 시도...");
                         
                         SetForegroundWindow(docWindow);
                         Thread.Sleep(200);
                         
                         Clipboard.SetText(text);
                         Thread.Sleep(100);
                         
                         // PostMessage 대신 SendMessage 사용하여 동기적으로 처리
                         SendMessage(docWindow, WM_KEYDOWN, (IntPtr)17, IntPtr.Zero); // Ctrl down
                         SendMessage(docWindow, WM_KEYDOWN, (IntPtr)86, IntPtr.Zero); // V down
                         SendMessage(docWindow, WM_KEYUP, (IntPtr)86, IntPtr.Zero);   // V up
                         SendMessage(docWindow, WM_KEYUP, (IntPtr)17, IntPtr.Zero);  // Ctrl up
                         Thread.Sleep(200);
                         
                         _logger?.LogInformation("개선된 키보드 메시지 방식 완료");
                         return true;
                     }
                     catch (Exception ex)
                     {
                         _logger?.LogWarning($"개선된 키보드 메시지 방식 실패: {ex.Message}");
                     }
                     
                     // 방법 2-4: WM_SETTEXT 직접 시도
                     try
                     {
                         _logger?.LogInformation("WM_SETTEXT 직접 시도...");
                         
                         var result = SendMessage(docWindow, WM_SETTEXT, IntPtr.Zero, text);
                         _logger?.LogInformation($"WM_SETTEXT 결과: {result}");
                         
                         if (result)
                         {
                             _logger?.LogInformation("WM_SETTEXT 방식 완료");
                             return true;
                         }
                     }
                     catch (Exception ex)
                     {
                         _logger?.LogWarning($"WM_SETTEXT 방식 실패: {ex.Message}");
                     }
                }

                // 3. 대체 방법: 한글 프로세스에 직접 WM_CHAR 전송
                try
                {
                    _logger?.LogInformation("한글 프로세스 직접 문자 전송 시도...");
                    
                    // 메인 윈도우를 활성화
                    SetForegroundWindow(mainWindow);
                    Thread.Sleep(200);
                    
                    // 텍스트를 한 글자씩 전송
                    foreach (char c in text)
                    {
                        PostMessage(mainWindow, WM_CHAR, (IntPtr)c, IntPtr.Zero);
                        Thread.Sleep(10);
                    }
                    
                    _logger?.LogInformation("프로세스 직접 문자 전송 완료");
                    return true;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"프로세스 직접 문자 전송 실패: {ex.Message}");
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"직접 문서 조작 중 오류: {ex.Message}");
                return false;
            }
        }

        // 한글의 문서 편집 영역 찾기
        private IntPtr FindHwpDocumentWindow(IntPtr parentWindow)
        {
            IntPtr bestCandidate = IntPtr.Zero;
            
            try
            {
                _logger?.LogInformation("한글 문서 편집 영역 탐색 중...");
                
                // 모든 자식 윈도우 탐색
                EnumChildWindows(parentWindow, (hWnd, lParam) =>
                {
                    var className = new StringBuilder(256);
                    GetClassName(hWnd, className, className.Capacity);
                    string classNameStr = className.ToString();
                    
                    // 윈도우 텍스트 확인
                    int textLength = GetWindowTextLength(hWnd);
                    string windowText = "";
                    if (textLength > 0)
                    {
                        var textBuilder = new StringBuilder(textLength + 1);
                        GetWindowText(hWnd, textBuilder, textBuilder.Capacity);
                        windowText = textBuilder.ToString();
                    }
                    
                    _logger?.LogDebug($"자식 윈도우 - 클래스: {classNameStr}, 텍스트: {windowText}, 가시성: {IsWindowVisible(hWnd)}");
                    
                    // 한글 문서 편집 영역으로 추정되는 윈도우들
                    if (IsWindowVisible(hWnd) && (
                        classNameStr.Contains("Hwp") ||
                        classNameStr.Contains("Edit") ||
                        classNameStr.Contains("Document") ||
                        classNameStr.Contains("RichEdit") ||
                        classNameStr.Contains("View") ||
                        windowText.Contains("문서") ||
                        (classNameStr.Length > 10 && !classNameStr.Contains("Button") && !classNameStr.Contains("Menu"))
                    ))
                    {
                        _logger?.LogInformation($"문서 편집 영역 후보 발견 - 클래스: {classNameStr}, 텍스트: {windowText}");
                        bestCandidate = hWnd;
                        
                        // 가장 적합한 후보를 찾으면 우선 선택
                        if (classNameStr.Contains("Hwp") && classNameStr.Contains("Edit"))
                        {
                            return false; // 탐색 중단
                        }
                    }
                    
                    return true; // 계속 탐색
                }, IntPtr.Zero);
                
                if (bestCandidate != IntPtr.Zero)
                {
                    _logger?.LogInformation($"최적의 문서 편집 영역 선택: {bestCandidate}");
                }
                else
                {
                    _logger?.LogWarning("적절한 문서 편집 영역을 찾을 수 없습니다.");
                }
                
                return bestCandidate;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"한글 문서 편집 영역 찾기 중 오류: {ex.Message}");
                return IntPtr.Zero;
            }
        }

         // 윈도우 영역 가져오기
         private RECT? GetWindowRect(IntPtr hWnd)
         {
             try
             {
                 RECT rect;
                 if (GetWindowRect(hWnd, out rect))
                 {
                     return rect;
                 }
                 return null;
             }
             catch
             {
                 return null;
             }
         }

         // 마우스 클릭 수행
         private void ClickWindow(int x, int y)
         {
             try
             {
                 // 마우스 커서를 해당 위치로 이동
                 SetCursorPos(x, y);
                 Thread.Sleep(50);
                 
                 // 마우스 클릭 수행
                 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                 Thread.Sleep(50);
                 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                 Thread.Sleep(50);
             }
             catch (Exception ex)
             {
                 _logger?.LogWarning($"마우스 클릭 실패: {ex.Message}");
             }
         }

        // HTML/XML 형식 텍스트 처리
        private string ProcessFormattedText(string text)
        {
            try
            {
                _logger?.LogInformation("텍스트 형식 분석 및 처리 중...");
                
                // HTML 형식인지 확인
                if (text.TrimStart().StartsWith("<") && text.TrimEnd().EndsWith(">"))
                {
                    _logger?.LogInformation("HTML/XML 형식 감지, 변환 중...");
                    return ConvertHtmlToPlainTextWithFormatting(text);
                }
                
                // XML 형식인지 확인 (다른 패턴)
                if (text.Contains("<") && text.Contains(">"))
                {
                    _logger?.LogInformation("부분 HTML/XML 태그 감지, 정리 중...");
                    return ConvertHtmlToPlainTextWithFormatting(text);
                }
                
                _logger?.LogInformation("일반 텍스트로 처리");
                return text;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"텍스트 처리 중 오류: {ex.Message}, 원본 텍스트 사용");
                return text;
            }
        }

        // HTML을 한글 호환 형식으로 변환
        private string ConvertHtmlToPlainTextWithFormatting(string html)
        {
            try
            {
                _logger?.LogInformation("HTML을 일반 텍스트로 변환 중...");
                
                string text = html;
                
                // 일반적인 HTML 태그 제거 및 변환
                text = System.Text.RegularExpressions.Regex.Replace(text, @"<br\s*/?>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                text = System.Text.RegularExpressions.Regex.Replace(text, @"<p[^>]*>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                text = System.Text.RegularExpressions.Regex.Replace(text, @"</p>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                text = System.Text.RegularExpressions.Regex.Replace(text, @"<div[^>]*>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                text = System.Text.RegularExpressions.Regex.Replace(text, @"</div>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                
                // 굵은 글씨, 기울임 등의 서식은 일단 태그 제거 (추후 한글 서식으로 변환 가능)
                text = System.Text.RegularExpressions.Regex.Replace(text, @"<[^>]+>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                
                // HTML 엔티티 디코딩
                text = System.Net.WebUtility.HtmlDecode(text);
                
                // 연속된 개행 정리
                text = System.Text.RegularExpressions.Regex.Replace(text, @"\n\s*\n", "\n\n");
                text = text.Trim();
                
                _logger?.LogInformation($"HTML 변환 완료, 길이: {text.Length}");
                return text;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"HTML 변환 실패: {ex.Message}, 원본 반환");
                return html;
            }
        }

        // HTML 클립보드를 통한 서식 적용
        private bool InsertFormattedTextViaClipboard(string text)
        {
            try
            {
                _logger?.LogInformation("클립보드를 통한 서식 적용 시작...");
                
                // HTML 형식인지 확인
                if (text.Contains("<") && text.Contains(">"))
                {
                    // HTML 형식으로 클립보드에 설정
                    _logger?.LogInformation("HTML 형식으로 클립보드 설정...");
                    
                    var dataObject = new System.Windows.Forms.DataObject();
                    dataObject.SetData(System.Windows.Forms.DataFormats.Html, text);
                    dataObject.SetData(System.Windows.Forms.DataFormats.Text, ConvertHtmlToPlainTextWithFormatting(text));
                    
                    Clipboard.SetDataObject(dataObject, true);
                    Thread.Sleep(200);
                }
                else
                {
                    // 일반 텍스트로 클립보드 설정
                    _logger?.LogInformation("일반 텍스트로 클립보드 설정...");
                    Clipboard.SetText(text);
                    Thread.Sleep(200);
                }
                
                // 한글 문서에 포커스 설정 후 붙여넣기
                var processes = Process.GetProcessesByName("Hwp");
                if (processes.Length > 0)
                {
                    IntPtr mainWindow = processes[0].MainWindowHandle;
                    SetForegroundWindow(mainWindow);
                    Thread.Sleep(300);
                    
                    // Ctrl+V로 붙여넣기
                    SendKeys.SendWait("^v");
                    Thread.Sleep(200);
                    
                    _logger?.LogInformation("클립보드 붙여넣기 완료");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"클립보드 서식 적용 실패: {ex.Message}");
                return false;
            }
        }

        // 한글 HAction을 통한 서식 적용
        private bool InsertFormattedTextViaHAction(string text)
        {
            try
            {
                _logger?.LogInformation("HAction을 통한 서식 적용 시작...");
                
                if (_hwpWrapper == null)
                {
                    _logger?.LogWarning("HwpWrapper가 null입니다.");
                    return false;
                }
                
                // 한글의 InsertText 액션 사용
                try
                {
                    _logger?.LogInformation("InsertText 액션 생성...");
                    
                    // HAction 생성
                    var action = _hwpWrapper.InvokeMethod("CreateAction", "InsertText");
                    if (action != null)
                    {
                        _logger?.LogInformation("HAction 생성 성공");
                        
                        // 파라미터 설정 (서식 포함 가능)
                        var paramSet = _hwpWrapper.InvokeMethod("CreateSet", "InsertText");
                        if (paramSet != null)
                        {
                            // 텍스트 설정
                            _hwpWrapper.InvokeMethod("HParameterAdd", paramSet, "Text", text);
                            
                            // 서식 관련 파라미터 설정 시도
                            try
                            {
                                // 폰트 설정 등 (선택적)
                                _hwpWrapper.InvokeMethod("HParameterAdd", paramSet, "FontFace", "맑은 고딕");
                                _hwpWrapper.InvokeMethod("HParameterAdd", paramSet, "FontSize", 10);
                            }
                            catch
                            {
                                // 서식 파라미터 설정 실패는 무시
                                _logger?.LogDebug("서식 파라미터 설정 생략");
                            }
                            
                            // 액션 실행
                            _hwpWrapper.InvokeMethod("HAction", action, paramSet);
                            _logger?.LogInformation("HAction 서식 적용 완료");
                            return true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"HAction 실행 실패: {ex.Message}");
                }
                
                // 대안: 직접 InsertText 호출
                try
                {
                    _logger?.LogInformation("직접 InsertText 호출...");
                    _hwpWrapper.InvokeMethod("InsertText", text);
                    _logger?.LogInformation("직접 InsertText 완료");
                    return true;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"직접 InsertText 실패: {ex.Message}");
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"HAction 서식 적용 실패: {ex.Message}");
                return false;
            }
        }

        // 새로운 Application 인스턴스 생성 시도
        private bool TryCreateNewApplicationInstance(string text)
        {
            try
            {
                _logger?.LogInformation("ProgID를 통한 새 Application 인스턴스 생성...");
                
                // 다양한 ProgID 시도
                string[] progIds = {
                    "HWPFrame.HwpObject",
                    "HWPFrame.HwpObject.1", 
                    "Hwp.Application",
                    "HwpApplication",
                    "HwpObject.Application"
                };

                foreach (string progId in progIds)
                {
                    try
                    {
                        _logger?.LogInformation($"ProgID '{progId}' 시도...");
                        Type? hwpType = Type.GetTypeFromProgID(progId);
                        if (hwpType != null)
                        {
                            object? hwpApp = Activator.CreateInstance(hwpType);
                            if (hwpApp != null)
                            {
                                _logger?.LogInformation($"새 Application 인스턴스 생성 성공: {progId}");
                                
                                var newWrapper = new HwpObjectWrapper(hwpApp, _logger);
                                if (newWrapper.InsertTextToDocument(text))
                                {
                                    _logger?.LogInformation("새 Application으로 텍스트 삽입 성공");
                                    return true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogDebug($"ProgID '{progId}' 실패: {ex.Message}");
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"새 Application 인스턴스 생성 실패: {ex.Message}");
                return false;
            }
        }

        public void Dispose()
        {
            try
            {
                if (_hwpWrapper != null)
                {
                    _hwpWrapper = null;
                }
                
                if (_hwpApp != null)
                {
                    Marshal.ReleaseComObject(_hwpApp);
                    _hwpApp = null;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"리소스 해제 중 오류 발생: {ex.Message}");
            }
        }

        private string[] FindHwpProgIDs()
        {
            var progIDs = new List<string>();
            try
            {
                _logger?.LogInformation("레지스트리에서 한글 ProgID 검색 중...");
                
                // HKEY_CLASSES_ROOT에서 HWP 관련 ProgID 찾기
                using (var hkcr = Registry.ClassesRoot)
                {
                    var possibleKeys = new[] { "HWPFrame", "Hwp", "HwpObject", "HwpApplication" };
                    
                    foreach (string keyName in hkcr.GetSubKeyNames())
                    {
                        if (possibleKeys.Any(k => keyName.StartsWith(k, StringComparison.OrdinalIgnoreCase)))
                        {
                            try
                            {
                                using (var subKey = hkcr.OpenSubKey(keyName))
                                {
                                    if (subKey != null)
                                    {
                                        var clsidKey = subKey.OpenSubKey("CLSID");
                                        if (clsidKey != null)
                                        {
                                            progIDs.Add(keyName);
                                            _logger?.LogInformation($"한글 ProgID 발견: {keyName}");
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // 액세스 권한 등의 이유로 실패할 수 있음
                            }
                        }
                    }
                }
                
                // 기본 ProgID들도 추가
                var defaultProgIDs = new[] { 
                    "HWPFrame.HwpObject", 
                    "HWPFrame.HwpObject.1", 
                    "Hwp.HwpObject", 
                    "HwpApplication",
                    "HwpObject",
                    "HwpObject.1"
                };
                
                foreach (var progId in defaultProgIDs)
                {
                    if (!progIDs.Contains(progId))
                    {
                        progIDs.Add(progId);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"레지스트리 검색 중 오류: {ex.Message}");
            }
            
            return progIDs.ToArray();
        }

        private bool TryRegisterHwpCom()
        {
            try
            {
                _logger?.LogInformation("한글 COM 라이브러리 수동 등록 시도...");
                
                // 한글 설치 경로에서 COM DLL 찾기
                string[] possibleDlls = {
                    @"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\HwpFrame.dll",
                    @"C:\Program Files\Hnc\Office 2020\HOffice110\Bin\HwpFrame.dll",
                    @"C:\Program Files (x86)\Hnc\Office 2022\HOffice120\Bin\HwpFrame.dll",
                    @"C:\Program Files\Hnc\Office 2022\HOffice120\Bin\HwpFrame.dll"
                };
                
                foreach (string dllPath in possibleDlls)
                {
                    if (File.Exists(dllPath))
                    {
                        _logger?.LogInformation($"한글 COM DLL 발견: {dllPath}");
                        
                        try
                        {
                            // regsvr32를 사용하여 DLL 등록
                            var processInfo = new ProcessStartInfo
                            {
                                FileName = "regsvr32",
                                Arguments = $"/s \"{dllPath}\"",
                                UseShellExecute = true,
                                Verb = "runas", // 관리자 권한으로 실행
                                WindowStyle = ProcessWindowStyle.Hidden
                            };
                            
                            using (var process = Process.Start(processInfo))
                            {
                                process?.WaitForExit(10000); // 10초 대기
                                _logger?.LogInformation($"COM DLL 등록 완료: {dllPath}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"COM DLL 등록 실패: {ex.Message}");
                        }
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"COM 라이브러리 등록 중 오류: {ex.Message}");
                return false;
            }
        }
    }
} 
