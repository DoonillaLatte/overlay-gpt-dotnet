using System;
using System.Windows;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Text.Json;

namespace overlay_gpt.Services
{
    public class MessageDispatcher
    {
        private static MessageDispatcher? _instance;
        private readonly ConcurrentQueue<Action> _messageQueue = new ConcurrentQueue<Action>();
        private readonly System.Windows.Threading.Dispatcher _uiDispatcher;
        private bool _isProcessing;

        public static MessageDispatcher Instance
        {
            get
            {
                if (_instance == null)
                {
                    throw new InvalidOperationException("MessageDispatcher가 초기화되지 않았습니다.");
                }
                return _instance;
            }
        }

        private MessageDispatcher(Window mainWindow)
        {
            _uiDispatcher = mainWindow.Dispatcher;
        }

        public static void Initialize(Window mainWindow)
        {
            if (_instance != null)
            {
                throw new InvalidOperationException("MessageDispatcher가 이미 초기화되었습니다.");
            }
            _instance = new MessageDispatcher(mainWindow);
        }

        public void DispatchToUI(Action action)
        {
            if (_uiDispatcher.CheckAccess())
            {
                action.Invoke();
            }
            else
            {
                _messageQueue.Enqueue(action);
                ProcessMessageQueue();
            }
        }

        public async Task DispatchToUIAsync(Action action)
        {
            if (_uiDispatcher.CheckAccess())
            {
                action.Invoke();
            }
            else
            {
                await _uiDispatcher.InvokeAsync(action);
            }
        }

        private void ProcessMessageQueue()
        {
            if (_isProcessing) return;

            _isProcessing = true;
            _uiDispatcher.BeginInvoke(new Action(() =>
            {
                while (_messageQueue.TryDequeue(out var action))
                {
                    try
                    {
                        action.Invoke();
                    }
                    catch (Exception ex)
                    {
                        // 예외 처리 로직 추가 가능
                        System.Diagnostics.Debug.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
                    }
                }
                _isProcessing = false;
            }));
        }

        public void DispatchJsonMessage(string messageType, JsonElement data)
        {
            DispatchToUI(() =>
            {
                // 여기에 메시지 타입에 따른 처리 로직을 추가할 수 있습니다
                // 예: 이벤트 발생, UI 업데이트 등
            });
        }
    }
} 