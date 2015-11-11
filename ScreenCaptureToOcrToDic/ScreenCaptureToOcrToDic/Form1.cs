using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using OpenCvSharp;
using SHDocVw;
using Tesseract;
using Utilities;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace ScreenCaptureToOcrToDic
{
    public partial class Form1 : Form
    {
        private static readonly InternetExplorer GoogleIe = new InternetExplorer();
        private static readonly InternetExplorer DaumIe = new InternetExplorer();
        private static readonly InternetExplorer NaverIe = new InternetExplorer();
        private static string _ocrText = string.Empty;
        private static int _targetWindowHandle;
        private readonly globalKeyboardHook _gkh = new globalKeyboardHook();

        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        private void Form1_Load(object sender, EventArgs e)
        {
            _gkh.HookedKeys.Add(Keys.Oemtilde);
            _gkh.HookedKeys.Add(Keys.D1);
            _gkh.HookedKeys.Add(Keys.D2);
            _gkh.HookedKeys.Add(Keys.D3);
            _gkh.KeyDown += GlobalKeyboardHookKeyDown;
            _gkh.KeyUp += GlobalKeyboardHookKeyUp;
        }

        private static void GlobalKeyboardHookKeyUp(object sender, KeyEventArgs e)
        {
            if (ModifierKeys == Keys.Control)
            {
                if (e.KeyCode == Keys.Oemtilde)
                {
                    SetTargetWindowHandle();
                }

                if (e.KeyCode == Keys.D1)
                {
                    // http://stackoverflow.com/questions/398882/invalidcastexception-rpc-e-cantcallout-ininputsynccall
                    // Entry(WebDicType.Naver);
                    var thread = new Thread(() => Entry(WebDicType.Naver));
                    thread.Start();
                    thread.Join(); //wait for the thread to finish
                }

                //if (e.KeyCode == Keys.D2)
                //{
                //    Entry(WebDicType.Google);
                //    var thread = new Thread(() => Entry(WebDicType.Google));
                //    thread.Start();
                //    thread.Join(); //wait for the thread to finish
                //}

                //if (e.KeyCode == Keys.D3)
                //{
                //    var thread = new Thread(() => Entry(WebDicType.Daum));
                //    thread.Start();
                //    thread.Join(); //wait for the thread to finish
                //}
            }

            e.Handled = true;
        }

        private static void GlobalKeyboardHookKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
        }

        // http://stackoverflow.com/questions/13547639/return-window-handle-by-its-name-title
        private static IntPtr WinGetHandle(string wName)
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

        [STAThread]
        private static void Entry(WebDicType webDicType)
        {
            Rectangle lpRect;
            GetWindowRect(new IntPtr(_targetWindowHandle), out lpRect);

            const string srcTestImagePath = "./test.png";
            const string dstTestImagePath = "./test1.png";

            const int titleHeight = 30;
            var bitmap = new Bitmap(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - titleHeight);
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.Left, lpRect.Top + titleHeight), new Point(0, 0),
                new Size(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - titleHeight));
            bitmap.Save(srcTestImagePath, ImageFormat.Png);
            graphics.Dispose();

            var src = Cv2.ImRead(srcTestImagePath, ImreadModes.GrayScale);
            var dst = new Mat(new[] {src.Width, src.Height}, MatType.CV_8U);
            Cv2.Threshold(src, dst, 254, 255, ThresholdTypes.Binary);
            Cv2.ImWrite(dstTestImagePath, dst);

            var engine = new TesseractEngine(@"./tessdata", "eng+en1", EngineMode.Default);
            var img = Pix.LoadFromFile(dstTestImagePath);
            var page = engine.Process(img);
            var newText = page.GetText();
            var origin = newText;
            newText = newText.Replace('\n', ' ');

            newText = newText.Replace('1', 'l');
            Console.WriteLine("before filter : " + newText);
            // newText = newText.Replace(@"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", " ");  // noise 필터
            newText = Regex.Replace(newText, @"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", "");
            newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z]+ ", ""); // 특수문자로 시작하는 문자열 필터링
            newText = Regex.Replace(newText, @"\b[b-zA-HJ-Z]\b", ""); // a,I를 뺀 외자 필터
            newText = Regex.Replace(newText, @"\*", ""); // 특수기호 필터
            newText = Regex.Replace(newText, @"\s+", " "); // 스페이스 여러개는 하나로 합침
            newText = newText.Trim();
            Console.WriteLine("after filter  : " + newText);

            _ocrText = newText;
            Console.WriteLine("Origin : \n" + origin);
            Console.WriteLine("\n\nOCR to Text : \n" + _ocrText);

            Console.WriteLine(0);

            switch (webDicType)
            {
                case WebDicType.Google:
                    var googleWebBrowser = (IWebBrowserApp) GoogleIe;
                    googleWebBrowser.Visible = true;
                    googleWebBrowser.Navigate("https://translate.google.com/?source=gtx_m#en/ko/" + _ocrText);
                    break;
                case WebDicType.Naver:
                    var naverWebBrowser = (IWebBrowserApp) NaverIe;
                    naverWebBrowser.Visible = true;
                    naverWebBrowser.Navigate(@"http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=" + _ocrText);
                    break;
                case WebDicType.Daum:
                    var daumWebBrowser = (IWebBrowserApp) DaumIe;
                    daumWebBrowser.Visible = true;
                    daumWebBrowser.Navigate(@"http://dic.daum.net/search.do?q=" + _ocrText + @"&t=word&dic=eng");
                    break;
            }
        }

        private static void SetTargetWindowHandle()
        {
            _targetWindowHandle = WindowFromPoint(MousePosition);
            Console.WriteLine("MousePosition x : {0}, y : {1}, windowHandle : {2}", MousePosition.X, MousePosition.Y,
                _targetWindowHandle);
        }

        private enum WebDicType
        {
            Google,
            Naver,
            Daum
        }
    }
}