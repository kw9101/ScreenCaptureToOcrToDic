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
            var ie = new InternetExplorer();
            var webBrowser = (IWebBrowserApp)ie;
            webBrowser.Visible = true;

            webBrowser.Navigate("http://www.google.com");

            Thread.Sleep(5000);
            webBrowser.Navigate("http://www.google.com");
            Thread.Sleep(5000);
            webBrowser.Quit();

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
            //aTimer = new Timer(3000);

            //// Hook up the Elapsed event for the timer.
            //aTimer.Elapsed += OnTimedEvent;
            //aTimer.Enabled = true;

            //Console.WriteLine("Press the Enter key to exit the program... ");
            //Console.ReadLine();
            //Console.WriteLine("Terminating the application...");
        }

        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            System.Drawing.Point mousePosition = Control.MousePosition;

            var hWnd = WindowFromPoint(mousePosition);
            Rectangle lpRect;
            var h = new System.IntPtr(hWnd);
            GetWindowRect(h, out lpRect);

            Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, hWnd);
            Console.WriteLine("Rect {0}, {1}", lpRect.Left, lpRect.Width);

            var bitmap = new Bitmap(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top);
            // 캡쳐할 화면의 사이즈 만큼의 영역사이즈 만큼의 bitmap을 생성해줍니다.

            var graphics = Graphics.FromImage(bitmap);
            // bitmap이미지를 기반으로 Grapics객체를 생성합니다. 
            // 이 클래스에 관해서는 C#부분에서 설명하도록 하겠습니다. (뭐 나도 공부하는거니깐 ㅠ)
            graphics.CopyFromScreen(new Point(lpRect.Left, lpRect.Top), new Point(0, 0), new Size(lpRect.Width - lpRect.Left, lpRect.Height - lpRect.Top));

            // 실제로 캡쳐를 하는 부분입니다. 그래픽에서 제공하는 CopyFromScreen 메서드를 사용합니다. 
            // 참고로 이 메서드는 " 픽셀의 사각형에 해당하는 색 데이터를 화면에서 
            // Graphics의 그리기 화면으로 bitblt(bit - block transfer)합니다.
            // 즉, 화면을 Graphics의 메모리에 다시 그려준다는 것이 되겠지요.

            //PointToScreen - 특정 클라이언트 지점의 위치를 화면 좌표로 계산합니다. 
            // form이 상속받은 control에 있는 함수입니다. 이 함수역시 C#에 올려놓도록 하겠슴당.
            //         ...... 

            var imageFactor = new ImageFactory();
            imageFactor.Load("./test.png");
            imageFactor.GaussianBlur(10);
            imageFactor.Save("./test2.png");

            var testImagePath = "./test2.png";

            // file을 저장하는 코드삽입
            // 이 Graphics와 아까 bitmap을 연결했으니.. bitmap에서 제공되는 저장메서드를 이용해서 
            //  원하는 위치에 따로 저장하면 끝입니다. ^^
            bitmap.Save(testImagePath, ImageFormat.Png);
            graphics.Dispose();

            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(testImagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        var text = page.GetText();
                        Console.WriteLine("OCR to Text : \n" + text);

                        //var webBrower = new WebBrowser();
                        var target = "http://endic.naver.com/popManager.nhn?sLn=kr&m=miniPopMain";

                        //webBrower.Navigate(target);
                        System.Diagnostics.Process.Start(target);
                    }
                }
            }
        }
    }
}
