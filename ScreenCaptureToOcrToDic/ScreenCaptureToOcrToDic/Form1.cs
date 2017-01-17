using System;
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
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        // http://stackoverflow.com/questions/13547639/return-window-handle-by-its-name-title

        //[STAThread]
        //private static void Entry(WebDicType webDicType)
        //{
        //    newText = newText.Replace('1', 'l');
        //    Console.WriteLine("before filter : " + newText);
        //    // newText = newText.Replace(@"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", " ");  // noise 필터
        //    newText = Regex.Replace(newText, @"([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+([a-zA-Z]*[^a-zA-Z\s]+[a-zA-Z]*)+", "");
        //    newText = Regex.Replace(newText, @"[^a-zA-Z\s][a-zA-Z]+ ", ""); // 특수문자로 시작하는 문자열 필터링
        //    newText = Regex.Replace(newText, @"\b[b-zA-HJ-Z]\b", ""); // a,I를 뺀 외자 필터
        //    newText = Regex.Replace(newText, @"\*", ""); // 특수기호 필터
        //    newText = Regex.Replace(newText, @"\s+", " "); // 스페이스 여러개는 하나로 합침
        //    newText = newText.Trim();
        //    Console.WriteLine("after filter  : " + newText);
        //}
    }
}