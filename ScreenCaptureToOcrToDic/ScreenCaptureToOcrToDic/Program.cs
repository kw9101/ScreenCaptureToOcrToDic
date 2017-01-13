using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using OpenCvSharp;
using Timer = System.Windows.Forms.Timer;
using Tesseract;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using SHDocVw;
using Utilities;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace ScreenCaptureToOcrToDic
{
    using System.Timers;

    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        // [STAThread]
        private static void Main()
        {
            // keyboard hook 초기화및 셋팅
            var gkh = new globalKeyboardHook();
            gkh.HookedKeys.Add(Keys.Up);
            gkh.HookedKeys.Add(Keys.Down);
            gkh.KeyDown += GkhKeyDown;
            gkh.KeyUp += GkhKeyUp;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        private static Timer aTimer;

        // [STAThread]
        //static void Main(string[] args)
        //{
        //    globalKeyboardHook gkh = new globalKeyboardHook();
        //    gkh.HookedKeys.Add(Keys.A);
        //    gkh.KeyDown += new KeyEventHandler(GkhKeyDown);
        //    gkh.KeyUp += new KeyEventHandler(GkhKeyUp);

        //    // 1초의 interval을 둔 timer 만들기
        //    aTimer = new Timer(5000);

        //    //// Hook up the Elapsed event for the timer.
        //    aTimer.Elapsed += OnTimedEvent;
        //    aTimer.Enabled = true;

        //    Console.WriteLine("Press the Enter key to exit the program... ");
        //    Console.ReadLine();
        //    Console.WriteLine("Terminating the application...");
        //}

        // 키를 누르고 있을 때 연속으로 눌리지 않게 하기 위한 꼼수.
        private static readonly Dictionary<Keys, bool> IsPresss = new Dictionary<Keys, bool>();

        private static void GkhKeyUp(object sender, KeyEventArgs e)
        {
            Console.WriteLine("Up\t" + e.KeyCode);

            if (IsPresss.ContainsKey(e.KeyCode))
            {
                IsPresss[e.KeyCode] = false;
            }

            e.Handled = true;
        }

        private static void GkhKeyDown(object sender, KeyEventArgs e)
        {
            bool isPress;
            if (IsPresss.TryGetValue(e.KeyCode, out isPress) == false)
            {
                IsPresss.Add(e.KeyCode, true);
                isPress = false;
            }
            else
            {
                IsPresss[e.KeyCode] = true;
            }

            if (isPress == false)
            {
                Console.WriteLine("Down\t" + e.KeyCode);

                switch (e.KeyCode)
                {
                    case Keys.Up:

                        break;
                    case Keys.Down:
                        break;
                    default:
                        break;
                }

                GetValue();
            }


            e.Handled = true;
        }

        private static readonly InternetExplorer GoogleIe = new InternetExplorer();
        private static readonly InternetExplorer DaumIe = new InternetExplorer();
        private static readonly InternetExplorer NaverIe = new InternetExplorer();
        private static string text = string.Empty;

        // http://stackoverflow.com/questions/13547639/return-window-handle-by-its-name-title
        public static IntPtr WinGetHandle(string wName)
        {
            var hWnd = IntPtr.Zero;
            foreach (var pList in Process.GetProcesses())
            {
                Console.WriteLine(pList.MainWindowTitle + " : " + pList.MainWindowHandle);
                if (pList.MainWindowTitle.Contains(wName))
                {
                    hWnd = pList.MainWindowHandle;
                }
            }
            return hWnd; //Should contain the handle but may be zero if the title doesn't match
        }

        //[STAThread]
        [SuppressMessage("ReSharper", "InvertIf")]
        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            GetValue();
        }

        private static void GetValue()
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            // hWnd = 67624;
            Rectangle lpRect;

            var h = new IntPtr(hWnd);
            GetWindowRect(h, out lpRect);

            // Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            //Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

            const string srcTestImagePath = "./test.png";
            const string dstTestImagePath = "./test1.png";

            var bitmap = new Bitmap(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - 30);
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.Left, lpRect.Top + 30), new Point(0, 0),
                new Size(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - 30));
            bitmap.Save(srcTestImagePath, ImageFormat.Png);
            graphics.Dispose();

            var src = Cv2.ImRead(srcTestImagePath, ImreadModes.GrayScale);
            var dst = new Mat(new int[] {src.Width, src.Height}, MatType.CV_8U);
            Cv2.Threshold(src, dst, 254, 255, ThresholdTypes.Binary);
            Cv2.ImWrite(dstTestImagePath, dst);

            using (var engine = new TesseractEngine(@"./tessdata", "eng+en1", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(dstTestImagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        var newText = page.GetText();
                        var origin = newText;
                        newText = newText.Replace('\n', ' ');

                        newText = newText.Replace('1', 'l');
                        Console.WriteLine("before filter : " + newText);
                        // newText = newText.Replace(@"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", " ");  // noise 필터
                        newText = Regex.Replace(newText, @"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s.]+[a-zA-Z]*)+",
                            "");
                        newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z]+ ", ""); // 특수문자로 시작하는 문자열 필터링
                        newText = Regex.Replace(newText, @"\b[b-zA-HJ-Z]\b", ""); // a,I를 뺀 외자 필터
                        newText = Regex.Replace(newText, @"\*", ""); // 특수기호 필터
                        newText = Regex.Replace(newText, @"\s+", " "); // 스페이스 여러개는 하나로 합침
                        newText = newText.Trim();
                        Console.WriteLine("after filter  : " + newText);

                        if (newText != text)
                        {
                            text = newText;
                            Console.WriteLine("Origin : \n" + origin);
                            Console.WriteLine("\n\nOCR to Text : \n" + text);

                            Console.WriteLine(0);

                            var googleWebBrowser = (IWebBrowserApp) GoogleIe;
                            googleWebBrowser.Visible = true;

                            var googleTarget = "https://translate.google.com/?source=gtx_m#en/ko/" + text;
                            googleWebBrowser.Navigate(googleTarget);

                            var daumWebBrowser = (IWebBrowserApp) DaumIe;
                            daumWebBrowser.Visible = true;
                            var daumTarget = @"http://dic.daum.net/search.do?q=" + text + @"&t=word&dic=eng";
                            daumWebBrowser.Navigate(daumTarget);

                            var naverWebBrowser = (IWebBrowserApp) NaverIe;
                            naverWebBrowser.Visible = true;
                            var naverTarget = @"http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=" + text;
                            naverWebBrowser.Navigate(naverTarget);

                            // var target = "http://translate.naver.com/#/en/ko/";
                            // var target = "http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=";
                            // var target = "http://endic.naver.com/popManager.nhn?sLn=kr&m=search&query=";
                            // target += text + @"&t=word&dic=eng";

                            //Console.WriteLine(target);
                            // webBrowser.Quit();
                        }
                    }
                }
            }
        }
    }
}
