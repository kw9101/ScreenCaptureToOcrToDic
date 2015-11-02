using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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

        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            var mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            Rectangle lpRect;
            var h = new System.IntPtr(hWnd);
            GetWindowRect(h, out lpRect);

            Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

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
            //imageFactor.Save(testImagePath2);

            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(testImagePath2))
                {
                    using (var page = engine.Process(img))
                    {
                        var newText = page.GetText();
                        newText = newText.Replace('\n', ' ');
                        newText = newText.Replace('*', ' ');
                        newText = newText.Replace('?', ' ');
                        newText = newText.Trim();

                        if (newText != text)
                        {
                            text = newText;

                            Console.WriteLine("OCR to Text : \n" + newText);

                            var webBrowser = (IWebBrowserApp)ie;
                            webBrowser.Visible = true;

                            // var target = "http://translate.naver.com/#/en/ko/";
                            var target = "http://endic.naver.com/search.nhn?sLn=kr&isOnlyViewEE=N&query=";
                            //var target = "http://endic.naver.com/popManager.nhn?sLn=kr&m=search&query=";
                            target += newText;

                            webBrowser.Navigate(target);
                            // webBrowser.Quit();
                        }
                    }
                }
            }
        }
    }
}
