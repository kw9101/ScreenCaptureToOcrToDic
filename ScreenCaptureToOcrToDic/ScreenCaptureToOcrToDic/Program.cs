using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OpenCvSharp;
using Tesseract;
using Utilities;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using KeyEventArgs = System.Windows.Forms.KeyEventArgs;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

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
            gkh.HookedKeys.Add(Keys.S);
            gkh.HookedKeys.Add(Keys.X);
            gkh.HookedKeys.Add(Keys.D);
            gkh.HookedKeys.Add(Keys.C);
            gkh.HookedKeys.Add(Keys.F);
            gkh.HookedKeys.Add(Keys.V);
            gkh.KeyDown += GkhKeyDown;
            gkh.KeyUp += GkhKeyUp;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        [DllImport("user32.dll")]
        private static extern int WindowFromPoint(Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern long GetWindowText(IntPtr hwnd, StringBuilder lpString, long cch);

        [DllImport("User32.dll")]
        private static extern IntPtr GetParent(IntPtr hwnd);

        private static void GkhKeyUp(object sender, KeyEventArgs e)
        {
            //Console.WriteLine("Up\t" + e.KeyCode);

            if (IsPresss.ContainsKey(e.KeyCode))
            {
                IsPresss[e.KeyCode] = false;
            }

            //e.Handled = false;
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
                switch (e.KeyCode)
                {
                    case Keys.A:
                        ToOcr(Translator.Google, 0.44f, 0.1f, 0.66f, true);
                        break;
                    case Keys.Z:
                        ToOcr(Translator.Google, 0.44f, 0.66f, 0.1f, true);
                        break;
                    case Keys.S:
                        ToOcr(Translator.Naver, 0.44f, 0.1f, 0.66f, true);
                        break;
                    case Keys.X:
                        ToOcr(Translator.Naver, 0.44f, 0.66f, 0.1f, true);
                        break;
                    case Keys.D:
                        ToOcr(Translator.Google, 0.44f, 0.1f, 0.66f, false);
                        break;
                    case Keys.C:
                        ToOcr(Translator.Google, 0.44f, 0.66f, 0.1f, false);
                        break;
                    case Keys.F:
                        ToOcr(Translator.Naver, 0.44f, 0.1f, 0.66f, false);
                        break;
                    case Keys.V:
                        ToOcr(Translator.Naver, 0.44f, 0.66f, 0.1f, false);
                        break;
                }
            }

            // Console.WriteLine("Down\t" + e.KeyCode);
            // e.Handled = false;
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

        private enum Translator
        {
            Google,
            Naver
        }

        private static void ToOcr(Translator translator, float crapLeftRightRatio, float crapTopRatio, float crapBottomRatio, bool isPostProcessing)
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            Rectangle lpRect;

            var h = new IntPtr(hWnd);

            var parentH = GetParent(h);

            var maxLength = GetWindowTextLength(parentH);
            var windowText = new StringBuilder("", maxLength + 5);
            GetWindowText(parentH, windowText, maxLength + 2);
            var caption = windowText.ToString();
            if (caption.Contains("CEMU") == false)
            {
                Console.WriteLine("Caption: " + caption);
                return;
            }

            GetWindowRect(h, out lpRect);

            // Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            //Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

            const string srcTestImagePath = "./test.png";
            const string dstTestImagePath = "./test1.png";

            var windowWidth = lpRect.Width - lpRect.X;
            var windowHeight = lpRect.Height - lpRect.Y;
            var crapLeftRight = (int)(windowWidth * crapLeftRightRatio);
            var crapTop = (int)(windowHeight * crapTopRatio);
            var crapBottom = (int)(windowHeight * crapBottomRatio);
            var crapWidth = windowWidth - crapLeftRight;
            var crapHeight = windowHeight - (crapTop + crapBottom);
            // Console.WriteLine("crapWidth: " + crapWidth + ", crapHeight: " + crapHeight + ", lpRect: " + lpRect);
            var bitmap = new Bitmap(crapWidth, crapHeight);
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.X + crapLeftRight / 2, lpRect.Y + crapTop), new Point(0, 0), new Size(crapWidth, crapHeight));
            bitmap.Save(srcTestImagePath, ImageFormat.Png);
            graphics.Dispose();

            //using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            //{
            //    using (var img = Pix.LoadFromFile(srcTestImagePath))
            //    {
            //        using (var page = engine.Process(img))
            //        {
            //            var newText = page.GetText();

            //            // url에 맞게 문자코드 수정
            //            newText = newText.Trim();
            //            if (string.IsNullOrEmpty(newText))
            //            {
            //                return;
            //            }

            //            newText = newText.Replace('\n', ' ');
            //            Console.WriteLine("1: " + newText);
            //            newText = newText.Replace(" ", "%20");
            //            newText = newText.Replace(@"""", "%22");
            //            // newText = @"%22" + newText + @"%22";

            //            //var naverTarget = "http://translate.naver.com/#/en/ko/" + newText;
            //            //Process.Start(naverTarget);

            //            var googleTarget = "https://translate.google.com/?source=gtx_m#en/ko/" + newText;
            //            Process.Start(googleTarget);
            //        }
            //    }
            //}

            if (isPostProcessing)
            {
                // 흑백 처리
                var src = Cv2.ImRead(srcTestImagePath, ImreadModes.GrayScale);
                var dst = new Mat(new[] { src.Width, src.Height }, MatType.CV_8U);
                Cv2.Threshold(src, dst, 150, 255, ThresholdTypes.Binary);
                Cv2.ImWrite(srcTestImagePath, dst);
            }

            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(srcTestImagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        var newText = page.GetText();

                        // url에 맞게 문자코드 수정
                        newText = newText.Trim();
                        if (string.IsNullOrEmpty(newText))
                        {
                            return;
                        }

                        newText = newText.Replace('\n', ' ');
                        Console.WriteLine(newText);
                        newText = newText.Replace(" ", "%20");
                        newText = newText.Replace(@"""", "%22");

                        switch (translator)
                        {
                            case Translator.Google:
                                newText = @"%22" + newText + @"%22";
                                var googleTarget = "https://translate.google.com/?source=gtx_m#en/ko/" + newText;
                                Process.Start(googleTarget);
                                break;
                            case Translator.Naver:
                                var naverTarget = "http://translate.naver.com/#/en/ko/" + newText;
                                Process.Start(naverTarget);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException(nameof(translator), translator, null);
                        }
                    }
                }
            }
        }
    }
}