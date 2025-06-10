using System.Collections.Generic;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Services
{
    public class ChatDataManager
    {
        private static ChatDataManager? _instance;
        private static readonly object _lock = new object();
        private readonly List<ChatData> _chatDataList;

        private ChatDataManager()
        {
            _chatDataList = new List<ChatData>();
        }

        public static ChatDataManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new ChatDataManager();
                    }
                }
                return _instance;
            }
        }

        public void AddChatData(ChatData chatData)
        {
            _chatDataList.Add(chatData);
        }

        public void AddChatDataRange(IEnumerable<ChatData> chatDataList)
        {
            _chatDataList.AddRange(chatDataList);
        }

        public List<ChatData> GetAllChatData()
        {
            return _chatDataList;
        }

        public ChatData? GetChatDataById(int chatId)
        {
            return _chatDataList.Find(x => x.ChatId == chatId);
        }
        
        public ChatData? GetChatDataByTimeStamp(string timeStamp)
        {
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] GetChatDataByTimeStamp 호출 - 찾는 타임스탬프: \"{timeStamp}\"");
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 현재 저장된 ChatData 목록:");
            foreach (var data in _chatDataList)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] - ChatId: {data.ChatId}, Timestamp: \"{data.GeneratedTimestamp}\"");
            }
            return _chatDataList.Find(x => x.GeneratedTimestamp.Equals(timeStamp, StringComparison.Ordinal));
        }

        public void RemoveChatData(int chatId)
        {
            var chatData = GetChatDataById(chatId);
            if (chatData != null)
            {
                _chatDataList.Remove(chatData);
            }
        }

        public void ClearAllChatData()
        {
            _chatDataList.Clear();
        }
    }
}
