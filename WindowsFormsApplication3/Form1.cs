using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenCvSharp;
using System.Threading;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll")]
        private static extern int FindWindow(string IpClassName, string IpWindowName);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(int hWnd1, int hWnd2, string Ipsz1, string Ipsz2);
        [DllImport("user32.dll")]
        public static extern int PostMessage(IntPtr hwnd, int wMsg, int wParam, IntPtr Iparam);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hDc, uint nFlags);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int  nRightRect, int nBottomRect);
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRgn(IntPtr hwnd, IntPtr hRgn);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);


        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_LBUTTONUP= 0x0202;
        const int BM_CLICK = 0x00F5;

        //프로그램 주소
        static IntPtr nhwnd = IntPtr.Zero;


        enum countImage : int
        {
            전투시작, 중보, victory, 상자, 전복, 다시하기
        }

        enum countItem : int
        {
            아이템확인, 판매, 획득
        }

        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

        public Form1()
        {
            InitializeComponent();
            MessageBox.Show("환영합니다.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                nhwnd = (IntPtr)FindWindow(null, "녹스 플레이어"); // 윈도우 창 제목

                if (!nhwnd.Equals(IntPtr.Zero))
                {
                    // 타이머 변수
                    timer = new System.Windows.Forms.Timer();
                    MoveWindow(nhwnd, 0, 0, 800, 600, true);
                    // 윈폼 타이머 사용
                    timer.Interval = 3000;//1000 * 30 + r.Next(0, 10); // 1분 + @
                    timer.Tick += new EventHandler(timer1_Tick);
                    timer.Start();
                }
                else
                {
                    MessageBox.Show("녹스 못찾겠어요");
                }
            }catch(Exception er)
            {
                MessageBox.Show("정상적으로 녹스플레이어를 찾지 못했습니다. " + er.Message.ToString() );
            }
            finally
            {

            }
        }

        //녹스플레이어 찾기
        private Bitmap PrintWindow2()
        {
            try
            {
                //찾은 플레이어를 바탕으로 Graphics 정보를 가져옵니다.
                Graphics Graphicsdata = Graphics.FromHwnd(nhwnd);

                //찾은 플레이어 창 크기 및 위치를 가져옵니다. 
                Rectangle rect = Rectangle.Round(Graphicsdata.VisibleClipBounds);

                //플레이어 창 크기 만큼의 비트맵을 선언해줍니다.
                Bitmap bmp = new Bitmap(rect.Width, rect.Height);

                //비트맵을 바탕으로 그래픽스 함수로 선언해줍니다.
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    //찾은 플레이어의 크기만큼 화면을 캡쳐합니다.
                    IntPtr hdc = g.GetHdc();
                    PrintWindow(nhwnd, hdc, 0x2);
                    g.ReleaseHdc(hdc);
                }
                
                // debug test
                //string path = @"c:\\NoxMacrotemp";
                //DirectoryInfo di = Directory.CreateDirectory(path);
                //bmp.Save(path+"\\7.bmp",System.Drawing.Imaging.ImageFormat.Bmp);
                return bmp;
            }
            catch (Exception e)
            {
                throw new Exception("녹스플레이어 찾지 못함 1" + e.Message.ToString());
            }
            finally
            {
                

            }

        }

        public System.Drawing.Point searchImg(Bitmap screen_img, Bitmap find_img)
        {
            //스크린 이미지
            Mat ScreenMat = new Mat();
            Mat FindMat = new Mat();
            Mat res = new Mat();
            using ( ScreenMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(screen_img))
            using ( FindMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(find_img))
            using( res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed))
            {
                double minval=0, maxval = 0;
                OpenCvSharp.Point minloc = new OpenCvSharp.Point(), maxloc =new OpenCvSharp.Point();

                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);

                Console.WriteLine("찾은 이미지 유사도 : " + maxval);
                System.Drawing.Point p = new System.Drawing.Point();
                
                if (maxval >= 0.8)
                {
                    textBox1.Text += "\n이미지 매칭 성공 ! 클릭준비중";
                    textBox1.Refresh();
                    Random rnd = new Random();
                    int x1 = rnd.Next(0, screen_img.Width / 10); 
                    int y1 = rnd.Next(0, screen_img.Height / 10);
                    p.X = maxloc.X + x1;
                    p.Y = maxloc.Y + y1;
                    //Inclick(maxloc.X + x1 , maxloc.Y + y1 , nhwnd);
                    return p;
                }

                return p;
            }
               
        }




        public void Inclick(int x, int y, IntPtr nhwnd)
        {
            
            if( nhwnd != IntPtr.Zero)
            {
                IntPtr lparam = new IntPtr(x | (y << 16));
                if (lparam != IntPtr.Zero)
                {
                    textBox2.Text = "해당이미지 클릭";
                    textBox2.Refresh();
                    Thread.Sleep(1000);
                   IntPtr nhwnd2 = FindWindowEx((int)nhwnd, 0, "Qt5QWindowIcon", "ScreenBoardClassWindow");

                    PostMessage(nhwnd2, WM_LBUTTONDOWN, 1, lparam);
                    Thread.Sleep(500);
                    PostMessage(nhwnd2, WM_LBUTTONUP, 0, lparam);
                    Thread.Sleep(600);
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Bitmap sell = new Bitmap(@"img\판매.PNG");
            Bitmap star5 = new Bitmap(@"img\5성희귀.PNG");
            //찾을 이미지 변수들
            Bitmap[] findImage =
            {
                new Bitmap(@"img\다시하기.PNG"),
                new Bitmap(@"img\전투시작.PNG"),
                new Bitmap(@"img\중보.PNG"),
                new Bitmap(@"img\victory.PNG"),
                new Bitmap(@"img\상자.PNG"),
                new Bitmap(@"img\전복아니오.PNG"),
                new Bitmap(@"img\전복.PNG")
            
            };

            Bitmap[] sellImage =
            {
                new Bitmap(@"img\아이템_확인.PNG"),
                new Bitmap(@"img\획득.PNG")
            };

            System.Drawing.Point p2 = System.Drawing.Point.Empty, p3= System.Drawing.Point.Empty;
            Random r = null;
            try
            {
                for (int i = 0; i < findImage.Length; i++)
                {
                    Thread.Sleep(2000);
                    p2 = new System.Drawing.Point();
                    p3 = new System.Drawing.Point();
                    r = new Random();

                    textBox2.Text = "";
                    textBox2.Refresh();
       
                    switch (i)
                    {
                        case (int)countImage.상자 + 1:
                            {
                                textBox1.Text = "";
                                textBox1.Text += "\n룬, 아이템 판별중";
                                textBox1.Refresh();
                                Thread.Sleep(500);
                                //5성희귀 판매 flag
                                bool sellFlag = starFlag_check.Checked;
                                Bitmap screen_img = PrintWindow2();
                                if (sellFlag)
                                {
                                    p2 = searchImg(screen_img, star5);
                                    if (!p2.IsEmpty)
                                    {
                                        p3 = searchImg(screen_img, sell);
                                        if (!p3.IsEmpty)
                                        {
                                            Inclick(p3.X, p3.Y, nhwnd);
                                        }
                                    }
                                }
                                else
                                {
                                    for (int j = 0; j < sellImage.Length; j++)
                                    {
                                        p2 = searchImg(screen_img, sellImage[j]);
                                        Inclick(p2.X, p2.Y, nhwnd);
                                    }
                                }
                                Thread.Sleep(600 + r.Next(0, 500));
                                break;
                            }


                        default:
                            {
                                textBox1.Text = "";
                                textBox1.Text += "\n이미지 확인" + i;
                                textBox1.Refresh();
                                Bitmap screen_img = PrintWindow2();
                                p2 = searchImg(screen_img, findImage[i]);
                                Inclick(p2.X, p2.Y, nhwnd);
                                Thread.Sleep(600 + r.Next(0, 500));
                                break;
                            }
                    }
                }
            }
            catch(Exception e2)
            {

            }
            finally
            {
                sell.Dispose();
                star5.Dispose();
                findImage = null;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("종료합니다.");
            timer.Stop();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
