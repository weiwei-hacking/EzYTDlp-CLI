using Ookii.Dialogs.WinForms;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EzYTDlpCLI
{
    class Program
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);

        private static string? ytDlpPath = string.Empty;
        private static string? ffmpegPath = string.Empty;
        private static bool isPlaylistDownload = false;
        private static bool isPlaylistLink = false;
        private static AppConfig Config = AppConfig.Load();
        private static ConsoleColor playlistColor = ConsoleColor.Red;

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDesktopWindow();

        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [STAThread]
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }

        private enum DownloadResult
        {
            Completed,
            BackToLinkSelection,
            ExitToMainMenu
        }

        private enum WorkflowResult
        {
            Continue,
            BackToMainMenu,
            Exit
        }

        static async Task MainAsync(string[] args)
        {
            Console.CursorVisible = false;
            Console.Title = "EzYTDlp CLI";
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;

            await Task.Run(() =>
            {
                Thread.Sleep(200);
                IntPtr handle = GetConsoleWindow();
                const int GWL_STYLE = -16;
                const int WS_MAXIMIZEBOX = 0x00010000;
                const int WS_SIZEBOX = 0x00040000;
                int style = GetWindowLong(handle, GWL_STYLE);
                style &= ~WS_MAXIMIZEBOX;
                style &= ~WS_SIZEBOX;

                if (IntPtr.Size == 8)
                    SetWindowLongPtr64(handle, GWL_STYLE, new IntPtr(style));
                else
                    SetWindowLong32(handle, GWL_STYLE, style);

                IntPtr hwnd = GetConsoleWindow();
                IntPtr desktop = GetDesktopWindow();
                GetWindowRect(desktop, out RECT screen);
                GetWindowRect(hwnd, out RECT rect);
                int consoleWidth = rect.Right - rect.Left;
                int consoleHeight = rect.Bottom - rect.Top;
                int x = (screen.Right - consoleWidth) / 2;
                int y = (screen.Bottom - consoleHeight) / 2;
                MoveWindow(hwnd, x, y, consoleWidth, consoleHeight, true);
            });

            ExtractEmbeddedResources();

            if (string.IsNullOrEmpty(ytDlpPath) || string.IsNullOrEmpty(ffmpegPath) ||
                !File.Exists(ytDlpPath) || !File.Exists(ffmpegPath))
            {
                MessageBox(IntPtr.Zero, "無法提取或找到 yt-dlp 或 ffmpeg 資源", "錯誤", 0x00000010U);
                return;
            }

            bool firstLaunch = true;

            while (true)
            {
                WorkflowResult result = WorkflowResult.Continue;

                if (firstLaunch && Config.SkipMainMenu)
                {
                    result = await StartWorkflow(true);
                    firstLaunch = false;
                }

                if (result == WorkflowResult.Exit)
                    break;

                Console.Clear();
                ShowLogoAsync(firstLaunch);
                firstLaunch = false;

                DisplayMenu(new[]
                {
                    "",
                    "[１] 開始使用",
                    "[２] 設定",
                    "",
                    "[０] 退出程式"
                }, alignOptions: true);

                ConsoleKey key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.D1)
                {
                    result = await StartWorkflow(false);
                }
                else if (key == ConsoleKey.D2)
                {
                    ShowSettings();
                    continue;
                }
                else if (key == ConsoleKey.D0)
                {
                    break;
                }
            }

            Config.Save();
        }

        private static int GetConsoleWidth(string s)
        {
            int width = 0;
            foreach (char c in s)
                width += c > 127 ? 2 : 1;
            return width;
        }

        private static void DisplayMenu(string[] lines, bool[]? coloredStates = null, int highlightIndex = -1, ConsoleColor highlightColor = ConsoleColor.Gray, bool alignOptions = false)
        {
            int width = Console.WindowWidth;
            int longestOption = alignOptions ? lines.Max(l => GetConsoleWidth(l)) : 0;

            for (int i = 0; i < lines.Length; i++)
            {
                Console.ResetColor();
                int left = alignOptions ? (width - longestOption) / 2 : (width - GetConsoleWidth(lines[i])) / 2;
                Console.SetCursorPosition(Math.Max(0, left), Console.CursorTop);

                if (i == highlightIndex)
                    Console.ForegroundColor = highlightColor;

                Console.WriteLine(lines[i]);
                Console.ResetColor();
            }
        }

        private static string? GetClipboardText()
        {
            string? text = null;
            Thread sta = new(() =>
            {
                try { text = Clipboard.GetText()?.Trim(); }
                catch { }
            });
            sta.SetApartmentState(ApartmentState.STA);
            sta.Start();
            sta.Join();
            return text;
        }

        private static bool IsYouTubeLink(string? link)
        {
            if (string.IsNullOrWhiteSpace(link)) return false;
            link = link.Trim();
            return link.Contains("youtube.com/") || link.Contains("youtu.be/");
        }

        private static async Task<WorkflowResult> StartWorkflow(bool playAnimation)
        {
            string? link = GetClipboardText() ?? "";
            bool confirmed = false;

            while (!confirmed)
            {
                Console.Clear();
                Console.WriteLine(CenterText("連結:"));
                Console.WriteLine(CenterText(link) + "\n");

                string[] options = new[]
                {
                    "[１] 確認開始",
                    "[２] 更改連結",
                    "[３] 更改連結(自動讀取剪貼簿)",
                    "",
                    "[０] 回到主選單"
                };
                DisplayMenu(options, alignOptions: true);

                ConsoleKey key = Console.ReadKey(true).Key;
                switch (key)
                {
                    case ConsoleKey.D1:
                        if (!IsYouTubeLink(link))
                            link = AskForLink();
                        else
                            confirmed = true;
                        break;

                    case ConsoleKey.D2:
                        link = AskForLink();
                        break;

                    case ConsoleKey.D3:
                        link = GetClipboardText() ?? "";
                        break;

                    case ConsoleKey.D0:
                        return WorkflowResult.BackToMainMenu;
                }
            }

            var downloadResult = await DownloadMain(link!);
            return downloadResult switch
            {
                DownloadResult.BackToLinkSelection => await StartWorkflow(false),
                DownloadResult.ExitToMainMenu => WorkflowResult.BackToMainMenu,
                _ => WorkflowResult.Continue
            };
        }

        private static void ShowLogoAsync(bool playAnimation)
        {
            string asciiArt = @"╔───────────────────────────────────────────────────────────────────────────────────╗
│                                                                                   │
│ ███████╗███████╗██╗   ██╗████████╗██████╗ ██╗     ██████╗      ██████╗██╗     ██╗ │
│ ██╔════╝╚══███╔╝╚██╗ ██╔╝╚══██╔══╝██╔══██╗██║     ██╔══██╗    ██╔════╝██║     ██║ │
│ █████╗    ███╔╝  ╚████╔╝    ██║   ██║  ██║██║     ██████╔╝    ██║     ██║     ██║ │
│ ██╔══╝   ███╔╝    ╚██╔╝     ██║   ██║  ██║██║     ██╔═══╝     ██║     ██║     ██║ │
│ ███████╗███████╗   ██║      ██║   ██████╔╝███████╗██║         ╚██████╗███████╗██║ │
│ ╚══════╝╚══════╝   ╚═╝      ╚═╝   ╚═════╝ ╚══════╝╚═╝          ╚═════╝╚══════╝╚═╝ │
│                                                                                   │
╚───────────────────────────────────────────────────────────────────────────────────╝
";
            string[] lines = asciiArt.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            int windowWidth = Console.WindowWidth;
            int startY = Math.Max(0, (Console.WindowHeight / 2) - (lines.Length / 2) - 4);

            foreach (string line in lines)
            {
                string trimmed = line.TrimEnd();
                int leftPadding = Math.Max(0, (windowWidth - trimmed.Length) / 2);
                Console.SetCursorPosition(leftPadding, startY++);
                Console.WriteLine(trimmed);
            }
        }

        private static string AskForLink()
        {
            Console.CursorVisible = true;
            Console.Write("\n請輸入 YouTube 連結: ");
            string? link = Console.ReadLine()?.Trim();
            Console.CursorVisible = false;
            return link ?? "";
        }

        private static async Task<DownloadResult> DownloadMain(string link)
        {
            isPlaylistLink = link.Contains("playlist") || link.Contains("list=");
            isPlaylistDownload = isPlaylistLink ? Config.DefaultPlaylist : true;
            playlistColor = isPlaylistDownload ? ConsoleColor.Green : ConsoleColor.Red;

            while (true)
            {
                Console.Clear();
                Console.WriteLine(CenterText("連結："));
                Console.WriteLine(CenterText(link) + "\n");

                string[] options = new[]
                {
                    "[１] 下載成影片",
                    "[２] 下載成音訊",
                    "[３] 下載整個清單",
                    "",
                    "[０] 回到連結選擇",
                };

                ConsoleColor actualColor = isPlaylistLink
                    ? (isPlaylistDownload ? ConsoleColor.Green : ConsoleColor.Red)
                    : ConsoleColor.DarkGray;

                DisplayMenu(options, highlightIndex: 2, highlightColor: actualColor, alignOptions: true);

                ConsoleKey key = Console.ReadKey(true).Key;
                switch (key)
                {
                    case ConsoleKey.D1:
                        await DownloadAsVideo(link);
                        break;
                    case ConsoleKey.D2:
                        await DownloadAsAudio(link);
                        break;
                    case ConsoleKey.D3:
                        if (isPlaylistLink)
                        {
                            isPlaylistDownload = !isPlaylistDownload;
                            playlistColor = isPlaylistDownload ? ConsoleColor.Green : ConsoleColor.Red;
                        }
                        break;
                    case ConsoleKey.D0:
                        return DownloadResult.BackToLinkSelection;
                }
            }
        }

        private static string CenterText(string text)
        {
            int width = Console.WindowWidth;
            int textWidth = GetConsoleWidth(text);
            int left = (width - textWidth) / 2;
            return new string(' ', Math.Max(0, left)) + text;
        }

        private static async Task RunYtDlpCommandRealtime(string args)
        {
            using Process proc = new()
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = ytDlpPath,
                    Arguments = args,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                },
                EnableRaisingEvents = true
            };

            proc.OutputDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) Console.WriteLine(e.Data); };
            proc.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(e.Data);
                    Console.ResetColor();
                }
            };

            Console.Clear();
            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            await proc.WaitForExitAsync();
        }

        private static async Task DownloadAsVideo(string link)
        {
            string? folder = SelectFolder(link);
            if (string.IsNullOrEmpty(folder)) return;

            string outputPath = Path.Combine(folder, "%(title)s.%(ext)s");
            string args = $"\"{link}\" -f \"bestvideo+bestaudio/best\" --output \"{outputPath}\"";

            if (Config.EmbedThumbnail) args += " --embed-thumbnail";
            if (!isPlaylistDownload && isPlaylistLink) args += " --no-playlist";

            await RunYtDlpCommandRealtime(args);
            ShowNotify(folder);
        }

        private static async Task DownloadAsAudio(string link)
        {
            string? folder = SelectFolder(link);
            if (string.IsNullOrEmpty(folder)) return;

            string outputPath = Path.Combine(folder, "%(title)s.%(ext)s");
            string args = $"\"{link}\" -x --audio-format mp3 --audio-quality 0 --output \"{outputPath}\"";

            if (Config.EmbedThumbnail) args += " --embed-thumbnail";
            if (!isPlaylistDownload && isPlaylistLink) args += " --no-playlist";

            await RunYtDlpCommandRealtime(args);
            ShowNotify(folder);
        }

        private static async Task<string> GetPlaylistTitle(string url)
        {
            try
            {
                using HttpClient client = new();
                string html = await client.GetStringAsync(url);
                var match = Regex.Match(html, @"<title>(.*?)<\/title>", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string title = match.Groups[1].Value
                        .Replace("- YouTube", "")
                        .Trim();
                    return Regex.Replace(title, @"[\\/:*?""<>|]", "_");
                }
            }
            catch { }
            return "Playlist";
        }

        private static string? SelectFolder(string link)
        {
            string? path = null;
            Thread sta = new(() =>
            {
                string basePath;

                if (Config.LockDownloadPath)
                {
                    basePath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                        "Downloads"
                    );
                }
                else
                {
                    using VistaFolderBrowserDialog dlg = new();
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return; // 使用者取消
                    basePath = dlg.SelectedPath;
                }

                path = basePath;

                if (isPlaylistLink && isPlaylistDownload)
                {
                    string playlistTitle = GetPlaylistTitle(link).GetAwaiter().GetResult();
                    path = Path.Combine(basePath, playlistTitle);
                    Directory.CreateDirectory(path);
                }
            });
            sta.SetApartmentState(ApartmentState.STA);
            sta.Start();
            sta.Join();

            return path;
        }


        private static void ShowNotify(string folder)
        {
            int result = MessageBox(IntPtr.Zero, "下載任務完成，要開啟資料夾嗎？", "完成通知", 0x00000004 | 0x00000040);
            if (result == 6)
                Process.Start("explorer.exe", folder);
        }

        private static void ExtractEmbeddedResources()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            ytDlpPath = Path.Combine(baseDir, "yt-dlp.exe");
            ffmpegPath = Path.Combine(baseDir, "ffmpeg.exe");
            ExtractResource("EzYTDlpCLI.yt-dlp.exe", ytDlpPath);
            ExtractResource("EzYTDlpCLI.ffmpeg.exe", ffmpegPath);
        }

        private static void ExtractResource(string name, string path)
        {
            if (File.Exists(path)) return;
            using Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name);
            if (stream == null) return;
            using FileStream fs = new(path, FileMode.Create, FileAccess.Write);
            stream.CopyTo(fs);
        }

        private static void ShowSettings()
        {
            while (true)
            {
                Console.Clear();

                string[] leftOptions = new[]
                {
                    "[１] 下載縮略圖",
                    "[２] 預設下載播放清單",
                    "[３] 鎖定下載路徑",
                    "[４] 跳過主選單",
                    "",
                    "[０] 回到主選單"
                };

                bool[] states = new[]
                {
                    Config.EmbedThumbnail,
                    Config.DefaultPlaylist,
                    Config.LockDownloadPath,
                    Config.SkipMainMenu
                };

                int[] totalLengths = new int[leftOptions.Length];
                for (int i = 0; i < leftOptions.Length; i++)
                {
                    string right = (i < states.Length) ? (states[i] ? "已啟用" : "已禁用") : "";
                    totalLengths[i] = GetConsoleWidth(leftOptions[i]) + 5 + GetConsoleWidth(right);
                }

                int maxTotalLength = totalLengths.Max();
                int leftPad = (Console.WindowWidth - maxTotalLength) / 2;
                int maxLeftLength = leftOptions.Take(states.Length).Max(l => GetConsoleWidth(l));

                for (int i = 0; i < leftOptions.Length; i++)
                {
                    Console.Write(new string(' ', Math.Max(0, leftPad)));
                    if (i < states.Length)
                    {
                        Console.Write(leftOptions[i]);
                        Console.Write(new string(' ', maxLeftLength - GetConsoleWidth(leftOptions[i]) + 5));
                        Console.ForegroundColor = states[i] ? ConsoleColor.Green : ConsoleColor.Red;
                        Console.WriteLine(states[i] ? "已啟用" : "已禁用");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine(leftOptions[i]);
                    }
                }

                ConsoleKey key = Console.ReadKey(true).Key;
                if (key == ConsoleKey.D0)
                    break;
                else if (key >= ConsoleKey.D1 && key <= ConsoleKey.D4)
                {
                    int index = key - ConsoleKey.D1;
                    switch (index)
                    {
                        case 0: Config.EmbedThumbnail = !Config.EmbedThumbnail; break;
                        case 1: Config.DefaultPlaylist = !Config.DefaultPlaylist; break;
                        case 2: Config.LockDownloadPath = !Config.LockDownloadPath; break;
                        case 3: Config.SkipMainMenu = !Config.SkipMainMenu; break;
                    }
                    Config.Save();
                }
            }
        }
    }

    public class AppConfig
    {
        public bool EmbedThumbnail { get; set; } = true;
        public bool DefaultPlaylist { get; set; } = false;
        public bool LockDownloadPath { get; set; } = true;
        public bool SkipMainMenu { get; set; } = true;
        public string? LastDownloadPath { get; set; } = "";

        private static readonly string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

        public static AppConfig Load()
        {
            if (!File.Exists(configPath)) return new AppConfig();
            try
            {
                string json = File.ReadAllText(configPath);
                return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
            }
            catch
            {
                return new AppConfig();
            }
        }

        public void Save()
        {
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configPath, json);
        }
    }
}
