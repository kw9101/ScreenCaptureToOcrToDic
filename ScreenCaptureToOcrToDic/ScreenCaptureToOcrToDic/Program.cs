using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Tesseract;
using Utilities;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace ScreenCaptureToOcrToDic
{
    internal class Program
    {
        // 키를 누르고 있을 때 연속으로 눌리지 않게 하기 위한 꼼수.
        private static readonly Dictionary<Keys, bool> IsPresss = new Dictionary<Keys, bool>();
        private static List<Tuple<Scalar, Scalar>> LetterColors = new List<Tuple<Scalar, Scalar>>();
        private static Dictionary<string, Rect> Areas = new Dictionary<string, Rect>();
        private static Dictionary<string, Tuple<string, Rect>> Behaviors = new Dictionary<string, Tuple<string, Rect>>();

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        private static void Main()
        {
            InitConfig();

            // keyboard hook 초기화및 셋팅
            var gkh = new globalKeyboardHook();

            foreach (var key in Behaviors.Keys)
            {
                var k = (Keys)char.ToUpper(key[0]);

                //gkh.HookedKeys.Add(k);
                Console.WriteLine(key + ": " + k);
            }

            // return;

            gkh.HookedKeys.Add(Keys.Q);
            gkh.HookedKeys.Add(Keys.A);
            gkh.HookedKeys.Add(Keys.Z);
            gkh.HookedKeys.Add(Keys.W);
            gkh.HookedKeys.Add(Keys.S);
            gkh.HookedKeys.Add(Keys.X);
            gkh.HookedKeys.Add(Keys.E);
            gkh.HookedKeys.Add(Keys.D);
            gkh.HookedKeys.Add(Keys.C);
            gkh.HookedKeys.Add(Keys.D1);
            gkh.HookedKeys.Add(Keys.D2);
            gkh.HookedKeys.Add(Keys.D3);
            gkh.HookedKeys.Add(Keys.D4);
            gkh.HookedKeys.Add(Keys.D7);
            gkh.HookedKeys.Add(Keys.D0);
            gkh.KeyDown += GkhKeyDown;
            gkh.KeyUp += GkhKeyUp;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void InitConfig()
        {
            // var configFilePath = @".\config.xml";
            var configFilePath = @"..\..\config.xml";
            if (File.Exists(configFilePath) == false)
            {
                configFilePath = @"..\..\config.xml";
                if (File.Exists(configFilePath) == false)
                {
                    Console.WriteLine("File is not exists.");
                }
            }

            var strXml = File.ReadAllText(configFilePath);

            var xml = new XmlDocument();
            xml.LoadXml(strXml);

            var colorNodes = xml.GetElementsByTagName("Color");
            foreach (XmlNode xn in colorNodes)
            {
                if (xn.Attributes == null)
                {
                    continue;
                }

                var minRgb = xn.Attributes["min"].Value.Split(',');
                var maxRgb = xn.Attributes["max"].Value.Split(',');

                var minr = Convert.ToInt32(minRgb[0]);
                var ming = Convert.ToInt32(minRgb[1]);
                var minb = Convert.ToInt32(minRgb[2]);

                var maxr = Convert.ToInt32(maxRgb[0]);
                var maxg = Convert.ToInt32(maxRgb[1]);
                var maxb = Convert.ToInt32(maxRgb[2]);

                LetterColors.Add(new Tuple<Scalar, Scalar>(new Scalar(minr, ming, minb), new Scalar(maxr, maxg, maxb)));
            }

            var areaNodes = xml.GetElementsByTagName("Area");
            foreach (XmlNode xn in areaNodes)
            {
                if (xn.Attributes == null)
                {
                    continue;
                }

                var name = xn.Attributes["name"].Value;
                var srcArea = xn.Attributes["area"].Value.Split(',');
                var area = new Rect(Convert.ToSingle(srcArea[0]), Convert.ToSingle(srcArea[1]), Convert.ToSingle(srcArea[2]), Convert.ToSingle(srcArea[3]));

                Areas.Add(name, area);
            }

            var behaviorNodes = xml.GetElementsByTagName("Behavior");
            foreach (XmlNode xn in behaviorNodes)
            {
                if (xn.Attributes == null)
                {
                    continue;
                }

                
                var key = xn.Attributes["key"].Value;
                var translate = xn.Attributes["translate"].Value;
                var area = xn.Attributes["area"].Value;

                Console.WriteLine(key);

                Behaviors.Add(key, new Tuple<string, Rect>(translate, Areas[area]));
            }
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
            // Console.WriteLine("Up\t" + e.KeyCode);

            if (IsPresss.ContainsKey(e.KeyCode))
            {
                IsPresss[e.KeyCode] = false;
            }

            //e.Handled = false;
        }

        private struct Rect
        {
            public Rect(float left, float right, float top, float bottom)
            {
                Left = left;
                Right = right;
                Top = top;
                Bottom = bottom;
            }

            public readonly float Left;
            public readonly float Right;
            public readonly float Top;
            public readonly float Bottom;
        }

        private static void GkhKeyDown(object sender, KeyEventArgs e)
        {
            bool isPress;
            if (IsPresss.TryGetValue(e.KeyCode, out isPress) == false)
            {
                IsPresss.Add(e.KeyCode, true);
                isPress = false;
            }

            var topRect = new Rect(0.22f, 0.22f, 0.1f, 0.66f);
            var middleRect = new Rect(0.22f, 0.22f, 0.39f, 0.39f);
            var bottomRect = new Rect(0.22f, 0.22f, 0.66f, 0.1f);

            //var LetterColors = new List<Tuple<Scalar, Scalar>>
            //{
            //    // 하양
            //    new Tuple<Scalar, Scalar>(new Scalar(200, 200, 200), new Scalar(255, 255, 255)),
            //    // 빨강
            //    new Tuple<Scalar, Scalar>(new Scalar(0, 50, 200), new Scalar(100, 150, 255)),
            //    // 초록
            //    new Tuple<Scalar, Scalar>(new Scalar(120, 200, 0), new Scalar(190, 255, 50)),
            //    // 파랑
            //    new Tuple<Scalar, Scalar>(new Scalar(180, 180, 0), new Scalar(255, 255, 80))
            //};

            IsPresss[e.KeyCode] = true;
            if (isPress == false)
            {
                switch (e.KeyCode)
                {
                    case Keys.D1:
                        if (ToOcr(Translator.Google, bottomRect, true, LetterColors))
                        {
                            break;
                        }

                        if (ToOcr(Translator.Google, topRect, true, LetterColors))
                        {
                            break;
                        }

                        ToOcr(Translator.Google, middleRect, true, LetterColors);

                        break;
                    case Keys.Q:
                        ToOcr(Translator.Google, topRect, true, LetterColors);
                        break;
                    case Keys.A:
                        ToOcr(Translator.Google, middleRect, true, LetterColors);
                        break;
                    case Keys.Z:
                        ToOcr(Translator.Google, bottomRect, true, LetterColors);
                        break;
                    case Keys.D2:
                        if (ToOcr(Translator.Google, bottomRect, false, LetterColors))
                        {
                            break;
                        }

                        if (ToOcr(Translator.Google, topRect, false, LetterColors))
                        {
                            break;
                        }

                        ToOcr(Translator.Google, middleRect, false, LetterColors);

                        break;
                    case Keys.W:
                        ToOcr(Translator.Google, topRect, false, LetterColors);
                        break;
                    case Keys.S:
                        ToOcr(Translator.Google, middleRect, false, LetterColors);
                        break;
                    case Keys.X:
                        ToOcr(Translator.Google, bottomRect, false, LetterColors);
                        break;
                    case Keys.D3:
                        if (ToOcr(Translator.Naver, bottomRect, false, LetterColors))
                        {
                            break;
                        }

                        if (ToOcr(Translator.Naver, topRect, false, LetterColors))
                        {
                            break;
                        }

                        ToOcr(Translator.Naver, middleRect, false, LetterColors);

                        break;
                    case Keys.E:
                        ToOcr(Translator.Naver, topRect, false, LetterColors);
                        break;
                    case Keys.D:
                        ToOcr(Translator.Naver, middleRect, false, LetterColors);
                        break;
                    case Keys.C:
                        ToOcr(Translator.Naver, bottomRect, false, LetterColors);
                        break;
                }
            }

            //Console.WriteLine("Down\t" + e.KeyCode);
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

        private static bool ToOcr(Translator translator, Rect crapRatio, bool isInQuotes, List<Tuple<Scalar, Scalar>> letterColors)
        {
            return ToOcr(translator, crapRatio.Left, crapRatio.Right, crapRatio.Top, crapRatio.Bottom, isInQuotes, letterColors);
        }

        private static bool ToOcr(Translator translator, float crapLeftRatio, float crapRightRatio, float crapTopRatio,
            float crapBottomRatio, bool isInQuotes, List<Tuple<Scalar, Scalar>> letterColors)
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
            if (caption.Contains("Cemu") == false)
            {
                Console.WriteLine("Caption: " + caption);
                return false;
            }

            GetWindowRect(h, out lpRect);

            const string srcTestImagePath = "./test.png";
            var windowWidth = lpRect.Width - lpRect.X;
            var windowHeight = lpRect.Height - lpRect.Y;

            var crapLeft = (int)(windowWidth * crapLeftRatio);
            var crapRight = (int)(windowWidth * crapRightRatio);
            var crapTop = (int)(windowHeight * crapTopRatio);
            var crapBottom = (int)(windowHeight * crapBottomRatio);

            var crapWidth = windowWidth - (crapLeft + crapRight);
            var crapHeight = windowHeight - (crapTop + crapBottom);

            var bitmap = new Bitmap(crapWidth, crapHeight);
            var graphics = Graphics.FromImage(bitmap);

            graphics.CopyFromScreen(new Point(lpRect.X + crapLeft, lpRect.Y + crapTop), new Point(0, 0),
                new Size(crapWidth, crapHeight));
            bitmap.Save(srcTestImagePath, ImageFormat.Png);
            graphics.Dispose();

            // 후처리
            var src = Cv2.ImRead(srcTestImagePath, ImreadModes.AnyColor);

            var letterFilter = new Mat(new[] { src.Width, src.Height }, MatType.CV_16U);
            var tempColorLetterfilter = new Mat(new[] { src.Width, src.Height }, MatType.CV_16U);

            Cv2.InRange(src, 0, 0, letterFilter);
            foreach (var letterColor in letterColors)
            {
                Cv2.InRange(src, letterColor.Item1, letterColor.Item2, tempColorLetterfilter);
                letterFilter += tempColorLetterfilter;
            }

            Cv2.ImWrite(srcTestImagePath, letterFilter);

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
                            return false;
                        }

                        newText = newText.Replace('\n', ' ');
                        Console.WriteLine(newText);
                        newText = newText.Replace(" ", "%20");
                        newText = newText.Replace(@"""", "%22");

                        if (isInQuotes)
                        {
                            newText = @"%22" + newText + @"%22";
                        }

                        switch (translator)
                        {
                            case Translator.Google:
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

            return true;
        }
    }
}