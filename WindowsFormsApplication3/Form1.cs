using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll")]
        private static extern int FindWindow(string IpClassName, string IpWindowName);
        [DllImport("user32.dll")]
        private static extern int FindWindowEx(int hWnd1, int hWnd2, string Ipsz1, string Ipsz2);
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hwnd, int wMsg, int wParam, int Iparam);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hDc, uint nFlags);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int  nRightRect, int nBottomRect);
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRgn(IntPtr hwnd, IntPtr hRgn);


        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_LBUTTONUP= 0x0202;
        const int BM_CLICK = 0x00F5;

        public Form1()
        {
            InitializeComponent();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
             IntPtr nhwnd = (IntPtr)FindWindow(null, "녹스 플레이어"); // 윈도우 창 제목

            if (!nhwnd.Equals(IntPtr.Zero)){
  
                PrintWindow2(nhwnd);
            }

            //if (nhwnd > 0)
            //{
            //    int hw2 = FindWindowEx(nhwnd, 0, "DirectUIHWND", "");
            //}
        }

        private static void PrintWindow2(IntPtr hwnd)
        {
            DateTime startTime = DateTime.Now;
            Rectangle rc = Rectangle.Empty;
            Graphics gfxWin = Graphics.FromHwnd(hwnd);
            rc = Rectangle.Round(gfxWin.VisibleClipBounds);
            Bitmap bmp = new Bitmap(
                rc.Width, rc.Height,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            
            Graphics gfxBmp = Graphics.FromImage(bmp);
            IntPtr hdcBitmap = gfxBmp.GetHdc();
            bool succeeded = PrintWindow(hwnd, hdcBitmap, 1);
            gfxBmp.ReleaseHdc(hdcBitmap);
            if (!succeeded)
            {
                gfxBmp.FillRectangle(
                        new SolidBrush(Color.Gray),
                        new Rectangle(Point.Empty, bmp.Size)
                    );
            }
            IntPtr hRgn = CreateRectRgn(0, 0, 0, 0);
            GetWindowRgn(hwnd, hRgn);
            Region region = Region.FromHrgn(hRgn);

            if (!region.IsEmpty(gfxBmp))
            {
                gfxBmp.ExcludeClip(region);
                gfxBmp.Clear(Color.Transparent);
            }
            gfxBmp.Dispose();
            Console.WriteLine((DateTime.Now - startTime).TotalMilliseconds);
            bmp.Save(@"c:\\test2.bmp",System.Drawing.Imaging.ImageFormat.Bmp);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
