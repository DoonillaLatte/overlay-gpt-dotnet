using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace overlay_gpt.Config
{
    public class AppConfig
    {
        [JsonPropertyName("openai_api_key")]
        public string OpenAiApiKey { get; set; } = string.Empty;

        [JsonPropertyName("app_version")]
        public string AppVersion { get; set; } = "1.0.0";

        [JsonPropertyName("flask_port")]
        public int FlaskPort { get; set; } = 5001;

        [JsonPropertyName("install_date")]
        public string InstallDate { get; set; } = string.Empty;
    }

    public static class ConfigManager
    {
        private static AppConfig? _config;
        private static readonly object _lock = new object();

        public static AppConfig Config
        {
            get
            {
                if (_config == null)
                {
                    lock (_lock)
                    {
                        if (_config == null)
                        {
                            LoadConfig();
                        }
                    }
                }
                return _config ?? new AppConfig();
            }
        }

        public static void LoadConfig()
        {
            try
            {
                // 1. 실행 파일과 같은 디렉토리에서 config.json 찾기
                string exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
                
                string[] configPaths = {
                    Path.Combine(exeDirectory, "config.json"),
                    Path.Combine(Directory.GetCurrentDirectory(), "config.json"),
                    Path.Combine(exeDirectory, "..", "config.json") // 상위 디렉토리
                };

                string? configContent = null;
                string? configPath = null;

                foreach (string path in configPaths)
                {
                    if (File.Exists(path))
                    {
                        configContent = File.ReadAllText(path);
                        configPath = path;
                        Console.WriteLine($"[ConfigManager] 설정 파일 로드: {path}");
                        break;
                    }
                }

                if (configContent != null)
                {
                    _config = JsonSerializer.Deserialize<AppConfig>(configContent, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });
                }
                else
                {
                    Console.WriteLine("[ConfigManager] config.json 파일을 찾을 수 없습니다.");
                    _config = new AppConfig();
                }

                // 2. 환경 변수에서 API 키 확인 (우선순위)
                string? envApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                if (!string.IsNullOrEmpty(envApiKey))
                {
                    _config.OpenAiApiKey = envApiKey;
                    Console.WriteLine("[ConfigManager] 환경 변수에서 OPENAI_API_KEY 로드됨");
                }

                // 3. API 키 검증
                if (string.IsNullOrEmpty(_config.OpenAiApiKey))
                {
                    Console.WriteLine("❌ [ConfigManager] OpenAI API 키를 찾을 수 없습니다!");
                    Console.WriteLine("설치 프로그램에서 API 키를 입력했는지 확인하거나");
                    Console.WriteLine("config.json 파일에 직접 설정해주세요.");
                }
                else if (IsApiKeyValid(_config.OpenAiApiKey))
                {
                    Console.WriteLine("✅ [ConfigManager] OpenAI API 키가 설정되었습니다.");
                }
                else
                {
                    Console.WriteLine("⚠️  [ConfigManager] OpenAI API 키 형식이 올바르지 않을 수 있습니다.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ [ConfigManager] 설정 로드 중 오류: {ex.Message}");
                _config = new AppConfig();
            }
        }

        public static bool IsApiKeyValid(string apiKey)
        {
            return !string.IsNullOrEmpty(apiKey) && 
                   apiKey.Length > 20 && 
                   apiKey.StartsWith("sk-", StringComparison.OrdinalIgnoreCase);
        }

        public static string GetOpenAiApiKey()
        {
            return Config.OpenAiApiKey;
        }

        public static int GetFlaskPort()
        {
            return Config.FlaskPort;
        }

        public static string GetFlaskUrl()
        {
            return $"http://localhost:{GetFlaskPort()}";
        }

        public static bool CreateSampleConfig(string filePath = "config.json")
        {
            try
            {
                var sampleConfig = new AppConfig
                {
                    OpenAiApiKey = "your-api-key-here",
                    AppVersion = "1.0.0",
                    FlaskPort = 5001,
                    InstallDate = DateTime.Now.ToString("yyyy-MM-dd")
                };

                string jsonString = JsonSerializer.Serialize(sampleConfig, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(filePath, jsonString);
                Console.WriteLine($"[ConfigManager] 샘플 설정 파일 생성됨: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ [ConfigManager] 샘플 설정 파일 생성 실패: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 설정을 다시 로드합니다.
        /// </summary>
        public static void ReloadConfig()
        {
            lock (_lock)
            {
                _config = null;
                LoadConfig();
            }
        }
    }
} 