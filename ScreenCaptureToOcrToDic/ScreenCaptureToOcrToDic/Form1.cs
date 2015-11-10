using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OpenCvSharp;
using SHDocVw;
using Tesseract;
using Utilities;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace ScreenCaptureToOcrToDic
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        readonly globalKeyboardHook _gkh = new globalKeyboardHook();

        static InternetExplorer googleIE = new InternetExplorer();
        static InternetExplorer daumIE = new InternetExplorer();
        static readonly InternetExplorer NaverIe = new InternetExplorer();
        private static string _ocrText = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _gkh.HookedKeys.Add(Keys.Q);
            _gkh.HookedKeys.Add(Keys.W);
            _gkh.KeyDown += GlobalKeyboardHookKeyDown;
            _gkh.KeyUp += GlobalKeyboardHookKeyUp;
        }

        static void GlobalKeyboardHookKeyUp(object sender, KeyEventArgs e)
        {
            if ((ModifierKeys == Keys.Control) && e.KeyCode == Keys.W)
            {
                SetTargetWindowHandle();
            }

            if (ModifierKeys == Keys.Control && e.KeyCode == Keys.Q)
            {
                Entry();
            }

            e.Handled = true;
        }

        static void GlobalKeyboardHookKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
        }

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

        private static int _targetWindowHandle = 0;

        //[STAThread]
        private static void Entry()
        {
            Rectangle lpRect;
            GetWindowRect(new IntPtr(_targetWindowHandle), out lpRect);

            const string srcTestImagePath = "./test.png";
            const string dstTestImagePath = "./test1.png";

            const int titleHeight = 30;
            var bitmap = new Bitmap(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - titleHeight);
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.Left, lpRect.Top + titleHeight), new Point(0, 0), new Size(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - titleHeight));
            bitmap.Save(srcTestImagePath, System.Drawing.Imaging.ImageFormat.Png);
            graphics.Dispose();

            var src = Cv2.ImRead(srcTestImagePath, ImreadModes.GrayScale);
            var dst = new Mat(new int[] { src.Width, src.Height }, MatType.CV_8U);
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
            newText = Regex.Replace(newText, @"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s.]+[a-zA-Z]*)+", "");
            newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z]+ ", "");   // 특수문자로 시작하는 문자열 필터링
            newText = Regex.Replace(newText, @"\b[b-zA-HJ-Z]\b", "");   // a,I를 뺀 외자 필터
            newText = Regex.Replace(newText, @"\*", ""); // 특수기호 필터
            newText = Regex.Replace(newText, @"\s+", " ");  // 스페이스 여러개는 하나로 합침
            newText = newText.Trim();
            Console.WriteLine("after filter  : " + newText);

            if (newText == _ocrText)
            {
                return;
            }

            _ocrText = newText;
            Console.WriteLine("Origin : \n" + origin);
            Console.WriteLine("\n\nOCR to Text : \n" + _ocrText);

            Console.WriteLine(0);

            //var googleWebBrowser = (IWebBrowserApp)googleIE;
            //googleWebBrowser.Visible = true;
            //var googleTarget = "https://translate.google.com/?source=gtx_m#en/ko/" + ocrText;
            //googleWebBrowser.Navigate(googleTarget);

            //var daumWebBrowser = (IWebBrowserApp)daumIE;
            //daumWebBrowser.Visible = true;
            //var daumTarget = @"http://dic.daum.net/search.do?q=" + ocrText + @"&t=word&dic=eng";
            //daumWebBrowser.Navigate(daumTarget);

            var naverWebBrowser = (IWebBrowserApp)NaverIe;
            naverWebBrowser.Visible = true;
            var naverTarget = @"http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=" + _ocrText;
            naverWebBrowser.Navigate(naverTarget);

            // var target = "http://translate.naver.com/#/en/ko/";
            // var target = "http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=";
            // var target = "http://endic.naver.com/popManager.nhn?sLn=kr&m=search&query=";
            // target += ocrText + @"&t=word&dic=eng";
        }

        private static void SetTargetWindowHandle()
        {
            _targetWindowHandle = WindowFromPoint(MousePosition);
            Console.WriteLine("MousePosition x : {0}, y : {1}, windowHandle : {2}", MousePosition.X, MousePosition.Y, _targetWindowHandle);
        }
    }
}