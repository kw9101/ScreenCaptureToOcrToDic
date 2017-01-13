using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Input;
using Tesseract;
using Utilities;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using KeyEventArgs = System.Windows.Forms.KeyEventArgs;

namespace ScreenCaptureToOcrToDic
{
    internal class Program
    {
        // 키를 누르고 있을 때 연속으로 눌리지 않게 하기 위한 꼼수.
        private static readonly Dictionary<Keys, bool> IsPresss = new Dictionary<Keys, bool>();

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        private static void Main()
        {
            // keyboard hook 초기화및 셋팅
            var gkh = new globalKeyboardHook();
           
            gkh.HookedKeys.Add(Keys.A);
            gkh.HookedKeys.Add(Keys.Z);
            gkh.HookedKeys.Add(Keys.LShiftKey);
            gkh.HookedKeys.Add(Keys.LControlKey);
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

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern long GetWindowText(IntPtr hwnd, StringBuilder lpString, long cch);

        [DllImport("User32.dll")]
        static extern IntPtr GetParent(IntPtr hwnd);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern long GetClassName(IntPtr hwnd, StringBuilder lpClassName, long nMaxCount);

        private static void GkhKeyUp(object sender, KeyEventArgs e)
        {
            //Console.WriteLine("Up\t" + e.KeyCode);

            if (IsPresss.ContainsKey(e.KeyCode))
            {
                IsPresss[e.KeyCode] = false;
            }

            e.Handled = false;
        }

        private static void GkhKeyDown(object sender, KeyEventArgs e)
        {
            bool isPress;
            if (IsPresss.TryGetValue(e.KeyCode, out isPress) == false)
            {
                IsPresss.Add(e.KeyCode, true);
                isPress = false;
            }

            IsPresss[e.KeyCode] = true;

            if (isPress == false)
            {
                bool isPressLShiftKey;
                IsPresss.TryGetValue(Keys.LControlKey, out isPressLShiftKey);
                bool isPressLControlKey;
                IsPresss.TryGetValue(Keys.LControlKey, out isPressLControlKey);

                if (isPressLShiftKey && isPressLControlKey)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.A:
                            GetValue(0.47f, 0.1f, 0.57f);
                            break;
                        case Keys.Z:
                            GetValue(0.47f, 0.57f, 0.09f);
                            break;
                    }
                }
            }

            // Console.WriteLine("Down\t" + e.KeyCode);
            e.Handled = false;
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

        private static void GetValue(float crapHRatio, float crapTopRatio, float crapBottomRatio)
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            // hWnd = 67624;
            Rectangle lpRect;

            var h = new IntPtr(hWnd);
            GetWindowRect(h, out lpRect);


            var maxLength = GetWindowTextLength(h);
            var windowText = new StringBuilder("", maxLength + 5);
            GetWindowText(h, windowText, maxLength + 2);

            var caption = windowText.ToString();
            if (caption != "Render")
            {
                Console.WriteLine("Caption: " + caption);
                return;
            }

            // Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            //Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

            const string srcTestImagePath = "./test.png";
            const string dstTestImagePath = "./test1.png";

            var crapH = (int) (lpRect.Width*crapHRatio);
            var crapTop = (int) (lpRect.Height*crapTopRatio);
            var crapBottom = (int) (lpRect.Height*crapBottomRatio);
            var bitmap = new Bitmap(lpRect.Width - lpRect.Left + crapH,
                lpRect.Height - lpRect.Top - (crapTop + crapBottom));
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.Left - crapH/2, lpRect.Top + crapTop), new Point(0, 0),
                new Size(lpRect.Width - lpRect.Left - crapH, lpRect.Height - lpRect.Top - (crapTop + crapBottom)));
            bitmap.Save(srcTestImagePath, ImageFormat.Png);
            graphics.Dispose();

            //var src = Cv2.ImRead(srcTestImagePath, ImreadModes.GrayScale);
            //var dst = new Mat(new int[] {src.Width, src.Height}, MatType.CV_8U);
            //Cv2.Threshold(src, dst, 254, 255, ThresholdTypes.Binary);
            //Cv2.ImWrite(dstTestImagePath, dst);

            // return;
            using (var engine = new TesseractEngine(@"./tessdata", "eng+en1", EngineMode.Default))
            {
                // using (var img = Pix.LoadFromFile(dstTestImagePath))
                using (var img = Pix.LoadFromFile(srcTestImagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        var newText = page.GetText();

                        // url에 맞게 문자코드 수정
                        newText = newText.Trim();
                        newText = newText.Replace('\n', ' ');
                        Console.WriteLine("Text: " + newText);
                        newText = newText.Replace(" ", "%20");
                        newText = newText.Replace(@"""", "%22");
                        newText = @"%22" + newText + @"%22";
                        // newText = newText.Replace('1', 'l');
                        //Console.WriteLine("before filter : " + newText);
                        // newText = newText.Replace(@"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", " ");  // noise 필터
                        //newText = Regex.Replace(newText, @"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s.]+[a-zA-Z]*)+", "");
                        //newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z]+ ", ""); // 특수문자로 시작하는 문자열 필터링
                        //newText = Regex.Replace(newText, @"\b[b-zA-HJ-Z]\b", ""); // a,I를 뺀 외자 필터
                        //newText = Regex.Replace(newText, @"\*", ""); // 특수기호 필터
                        //newText = Regex.Replace(newText, @"\s+", " "); // 스페이스 여러개는 하나로 합침
                        //newText = newText.Trim();
                        //Console.WriteLine("after filter  : " + newText);

                        var naverTarget = "http://translate.naver.com/#/en/ko/" + newText;
                        Process.Start(naverTarget);

                        var googleTarget = "https://translate.google.com/?source=gtx_m#en/ko/" + newText;
                        Process.Start(googleTarget);
                    }
                }
            }
        }
    }
}