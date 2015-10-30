using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace ScreenCaptureToOcrToDic
{
    using System.Timers;

    class Program
    {
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(Point lpPoint);

        private static Timer aTimer;

        static void Main(string[] args)
        {
            // 1초의 interval을 둔 timer 만들기
            aTimer = new Timer(1000);

            // Hook up the Elapsed event for the timer.
            aTimer.Elapsed += OnTimedEvent;
            aTimer.Enabled = true;

            Console.WriteLine("Press the Enter key to exit the program... ");
            Console.ReadLine();
            Console.WriteLine("Terminating the application...");
        }

        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            System.Drawing.Point mousePosition = Control.MousePosition;

            var windowHandle = WindowFromPoint(mousePosition);
            Console.WriteLine("Time : {0} > MousePosition x : {1}, y : {2}, windowHandle : {3}", e.SignalTime, mousePosition.X, mousePosition.Y, windowHandle);
        }
    }
}
