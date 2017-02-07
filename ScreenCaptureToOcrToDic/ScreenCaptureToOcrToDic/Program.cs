using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using OpenCvSharp;
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

        private static readonly Dictionary<string, Tuple<Scalar, Scalar>> LetterColors = new Dictionary<string, Tuple<Scalar, Scalar>>();

        private static readonly Dictionary<string, Rect> Areas = new Dictionary<string, Rect>();

        private static readonly Dictionary<Keys, Func<bool>> ToOcrs = new Dictionary<Keys, Func<bool>>();

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        private static void Main()
        {
            InitConfig(ToOcrs);

            // keyboard hook 초기화및 셋팅
            var gkh = new globalKeyboardHook();

            foreach (var key in ToOcrs.Keys)
            {
                gkh.HookedKeys.Add(key);
            }

            gkh.KeyDown += GkhKeyDown;
            gkh.KeyUp += GkhKeyUp;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void InitConfig(Dictionary<Keys, Func<bool>> toOcrs)
        {
            var configFilePath = @".\config.xml";
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

                var name = xn.Attributes["name"].Value;
                var minRgb = xn.Attributes["min"].Value.Split(',');
                var maxRgb = xn.Attributes["max"].Value.Split(',');

                var minr = Convert.ToInt32(minRgb[0]);
                var ming = Convert.ToInt32(minRgb[1]);
                var minb = Convert.ToInt32(minRgb[2]);

                var maxr = Convert.ToInt32(maxRgb[0]);
                var maxg = Convert.ToInt32(maxRgb[1]);
                var maxb = Convert.ToInt32(maxRgb[2]);

                LetterColors.Add(name,
                    new Tuple<Scalar, Scalar>(new Scalar(minr, ming, minb), new Scalar(maxr, maxg, maxb)));
            }

            var areaNodes = xml.GetElementsByTagName("Area");
            foreach (XmlNode xn in areaNodes)
            {
                if (xn.Attributes == null)
                {
                    continue;
                }

                var name = xn.Attributes["name"].Value;
                var srcArea = xn.Attributes["area"].Value.Split(',').Select(p => p.Trim()).ToList();
                var area = new Rect(Convert.ToSingle(srcArea[0]), Convert.ToSingle(srcArea[1]),
                    Convert.ToSingle(srcArea[2]), Convert.ToSingle(srcArea[3]));

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

                var translator = Translator.Google;
                switch (translate.ToLower())
                {
                    case "google":
                        translator = Translator.Google;
                        break;
                    case "googleinquotes":
                        translator = Translator.GoogleInQuotes;
                        break;
                    case "naver":
                        translator = Translator.Naver;
                        break;
                }

                var srcAreas = xn.Attributes["areas"].Value.Split(',').Select(p => p.Trim()).ToList();
                var areas = srcAreas.Select(srcArea => Areas[srcArea]).ToList();

                var srcColors = xn.Attributes["colors"].Value.Split(',').Select(p => p.Trim()).ToList();
                var colors = new List<Tuple<Scalar, Scalar>>();
                colors.AddRange(srcColors[0].ToLower() == "all"
                    ? LetterColors.Values
                    : srcColors.Select(srcColor => LetterColors[srcColor]));

                toOcrs.Add((Keys)char.ToUpper(key[0]), () => ToOcr(translator, areas, colors));
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

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll")]
        private static extern bool ShowWindow(IntPtr handle, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private static void GkhKeyUp(object sender, KeyEventArgs e)
        {
            // Console.WriteLine("Up\t" + e.KeyCode);

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
                ToOcrs[e.KeyCode]();
            }
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

        private static bool ToOcr(Translator translator, IEnumerable<Rect> crapRatios,
            IEnumerable<Tuple<Scalar, Scalar>> letterColors)
        {
            foreach (var crapRatio in crapRatios)
            {
                return ToOcr(translator, crapRatio.Left, crapRatio.Right, crapRatio.Top, crapRatio.Bottom, letterColors);
            }

            return false;
        }

        private static bool ToOcr(Translator translator, float crapLeftRatio, float crapRightRatio, float crapTopRatio,
            float crapBottomRatio, IEnumerable<Tuple<Scalar, Scalar>> letterColors)
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            Rectangle lpRect;

            var h = new IntPtr(hWnd);

            var parentH = GetParent(h);

            var maxLength = GetWindowTextLength(parentH);
            var windowText = new StringBuilder("", maxLength + 5);
            GetWindowText(parentH, windowText, maxLength + 2);
            // GetWindowText(h, windowText, maxLength + 2);
            var caption = windowText.ToString();
            //if (caption.Contains("Cemu") == false)
            //{
                Console.WriteLine("Caption: " + h + ", " + caption);
            //    return false;
            //}

            GetWindowRect(h, out lpRect);

            const string srcTestImagePath = "./test.png";
            var windowWidth = lpRect.Width - lpRect.X;
            var windowHeight = lpRect.Height - lpRect.Y;

            var crapLeft = (int) (windowWidth*crapLeftRatio);
            var crapRight = (int) (windowWidth*crapRightRatio);
            var crapTop = (int) (windowHeight*crapTopRatio);
            var crapBottom = (int) (windowHeight*crapBottomRatio);

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

            var letterFilter = new Mat(new[] {src.Width, src.Height}, MatType.CV_16U);
            var tempColorLetterfilter = new Mat(new[] {src.Width, src.Height}, MatType.CV_16U);

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

                        string target;
                        switch (translator)
                        {
                            case Translator.Google:
                                target = "https://translate.google.com/?source=gtx_m#en/ko/" + newText;
                                break;
                            case Translator.GoogleInQuotes:
                                target = "https://translate.google.com/?source=gtx_m#en/ko/" + @"%22" + newText + @"%22";
                                break;
                            case Translator.Naver:
                                target = "http://translate.naver.com/#/en/ko/" + newText;
                                break;
                            default:
                                throw new ArgumentOutOfRangeException(nameof(translator), translator, null);
                        }

                        Process.Start(target);

                        //Process[] processes = Process.GetProcessesByName("Shantae and the Pirate's Curse");

                        //foreach (Process p in processes)
                        //{
                        //    ShowWindow(p.MainWindowHandle, 9);
                        //    SetForegroundWindow(p.MainWindowHandle);
                        //}

                        // Process.GetProcessesByName("Shantae and the Pirate's Curse");
                        //var hWnd2 = FindWindow(null, "Shantae and the Pirate's Curse");
                        //Console.WriteLine("hwnd2: " + hWnd2);

                        //ShowWindow(hWnd2, 9);
                        //SetForegroundWindow(hWnd2);
                    }
                }
            }

            return true;
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

        private enum Translator
        {
            Google,
            GoogleInQuotes,
            Naver
        }
    }
}