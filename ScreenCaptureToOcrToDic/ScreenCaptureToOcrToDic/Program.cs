using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using ImageProcessor;
using ImageProcessor.Imaging;
using Timer = System.Windows.Forms.Timer;
using Tesseract;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using SHDocVw;

namespace ScreenCaptureToOcrToDic
{
    using System.Timers;

    class Program
    {
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        private static Timer aTimer;
        // [STAThread]
        static void Main(string[] args)
        {
            
            // http://imageprocessor.org/imageprocessor/imagefactory/resize/

            //var imageFactor = new ImageFactory();
            //imageFactor.Load("./test.png");
            //var size = imageFactor.Image.Size;
            //size.Width *= 2;
            //size.Height *= 2;
            //imageFactor.Resize(size);

            ////// imageFactor.DetectEdges()
            //imageFactor.GaussianBlur(3);
            //////imageFactor.Saturation(0);

            ////// imageFactor.GaussianSharpen(3);
            //imageFactor.Save("./test2.png");

            //var testImagePath = "./test2.png";

            //using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            //{
            //    using (var img = Pix.LoadFromFile(testImagePath))
            //    {
            //        using (var page = engine.Process(img))
            //        {
            //            var text = page.GetText();
            //            Console.WriteLine("OCR to Text :\n" + text);
            //        }
            //    }
            //}

            // 1초의 interval을 둔 timer 만들기
            aTimer = new Timer(3000);

            //// Hook up the Elapsed event for the timer.
            aTimer.Elapsed += OnTimedEvent;
            aTimer.Enabled = true;

            Console.WriteLine("Press the Enter key to exit the program... ");
            Console.ReadLine();
            Console.WriteLine("Terminating the application...");
        }

        static InternetExplorer ie = new InternetExplorer();
        private static string text = string.Empty;

        //[STAThread]
        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            hWnd = 133776;
            Rectangle lpRect;
            
            var h = new System.IntPtr(hWnd);
            GetWindowRect(h, out lpRect);

            Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            //Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

            var testImagePath = "./test.png";
            var testImagePath2 = "./test11.png";

            var bitmap = new Bitmap(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - 25);
            var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(new Point(lpRect.Left, lpRect.Top + 25), new Point(0, 0), new Size(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top - 25));
            bitmap.Save(testImagePath, ImageFormat.Png);
            graphics.Dispose();

            //var imageFactor = new ImageFactory();
            //imageFactor.Load(testImagePath);
            //var size = imageFactor.Image.Size;
            //size.Width *= 2;
            //size.Height *= 2;
            //imageFactor.Resize(size);
            //imageFactor.GaussianBlur(1);
            //// imageFactor.GaussianSharpen(1);
            //imageFactor.Save(testImagePath2);

            using (var engine = new TesseractEngine(@"./tessdata", "eng+en1", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(testImagePath))
                {
                    using (var page = engine.Process(img))
                    {
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
                        //newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z][^a-zA-Z\s]", "");
                        //newText = Regex.Replace(newText, @"([a-zA-Z]+)7", "$1?");
                        //newText = Regex.Replace(newText, @"1([a-zA-Z]+)", "l$1");
                        //newText = Regex.Replace(newText, @"([a-zA-Z]+)1", "$11");

                        //newText = Regex.Replace(newText, "[^a-zA-Z -'!.][a-zA-Z][^a-zA-Z -'!.]", "");
                        //newText = Regex.Replace(newText, "[^a-zA-Z -'!.]", "");

                        //newText = Regex.Replace(newText, @"\s+", " ");
                        //newText = Regex.Replace(newText, "[^a-zA-Z -'!.]{2,}", "");
                        //newText = Regex.Replace(newText, " [b-zA-HJ-Z -'!.] ", "");

                        if (newText != text)
                        {
                            text = newText;
                            Console.WriteLine("Origin : \n" + origin);
                            Console.WriteLine("\n\nOCR to Text : \n" + text);

                            Console.WriteLine(0);
                            var webBrowser = (IWebBrowserApp)ie;
                            webBrowser.Visible = true;

                            // var target = @"http://dic.daum.net/search.do?q=" + text + @"&t=word&dic=eng";
                            var target = "https://translate.google.com/?source=gtx_m#en/ko/" + text;
                            // var target = "http://translate.naver.com/#/en/ko/";
                            // var target = "http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=";
                            // var target = "http://endic.naver.com/popManager.nhn?sLn=kr&m=search&query=";
                            // target += text + @"&t=word&dic=eng";

                            //Console.WriteLine(target);
                            webBrowser.Navigate(target);
                            // webBrowser.Quit();
                        }
                    }
                }
            }
        }
    }
}
