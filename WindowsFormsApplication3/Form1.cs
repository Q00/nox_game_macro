using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenCvSharp;
using System.Threading;
using System.Threading.Tasks;

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
        IntPtr nhwnd = IntPtr.Zero;


        enum countImage : int
        {
            전투시작, 중보, victory, 상자, 전복, 다시하기
        }

        enum countItem : int
        {
            아이템확인, 판매, 획득
        }

        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

        System.Drawing.Point p2 = System.Drawing.Point.Empty;

        public Form1()
        {
            InitializeComponent();
            MessageBox.Show("환영합니다.");
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            nhwnd = (IntPtr)FindWindow(null, "녹스 플레이어"); // 윈도우 창 제목

            try
            {
                if (!nhwnd.Equals(IntPtr.Zero))
                {
                    // 타이머 변수
                    timer = new System.Windows.Forms.Timer();
                    MoveWindow(nhwnd, 0, 0, 800, 600, true);
                    timer1_Tick(sender, e);
                    // 윈폼 타이머 사용
                    timer.Interval = 1000*60;//1000 * 30 + r.Next(0, 10); // 1분 + @
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
                DateTime startTime = DateTime.Now;
                Rectangle rc = Rectangle.Empty;
                Graphics gfxWin = Graphics.FromHwnd(nhwnd);
                rc = Rectangle.Round(gfxWin.VisibleClipBounds);
                Bitmap bmp = new Bitmap(
                    rc.Width, rc.Height
                );

                Graphics gfxBmp = Graphics.FromImage(bmp);
                IntPtr hdcBitmap = gfxBmp.GetHdc();
                bool succeeded = PrintWindow(nhwnd, hdcBitmap, 1);
                gfxBmp.ReleaseHdc(hdcBitmap);
                if (!succeeded)
                {
                    gfxBmp.FillRectangle(
                            new SolidBrush(Color.Gray),
                            new Rectangle(System.Drawing.Point.Empty, bmp.Size)
                        );
                }
                IntPtr hRgn = CreateRectRgn(0, 0, 0, 0);
                GetWindowRgn(nhwnd, hRgn);
                Region region = Region.FromHrgn(hRgn);

                if (!region.IsEmpty(gfxBmp))
                {
                    gfxBmp.ExcludeClip(region);
                    gfxBmp.Clear(Color.Transparent);
                }
                gfxBmp.Dispose();
                Console.WriteLine((DateTime.Now - startTime).TotalMilliseconds);
                //// debug test
                string path = @"c:\\NoxMacrotemp";
                System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(path);
                bmp.Save(path + "\\7.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
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
            System.Drawing.Point p = new System.Drawing.Point();
            Random rnd = new Random();
            Mat ScreenMat = new Mat();
            Mat FindMat = new Mat();
            Mat res = new Mat();
            try
            {
                //스크린 이미지 선언
                using (ScreenMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(screen_img))
                //찾을 이미지 선언
                using (FindMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(find_img))
                using (res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed))
                {
                    double minval = 0, maxval = 0;
                    OpenCvSharp.Point minloc = new OpenCvSharp.Point(), maxloc = new OpenCvSharp.Point();

                    Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);

                    Console.WriteLine("찾은 이미지 유사도 : " + maxval);
                    

                    if (maxval >= 0.8)
                    {
                        textBox1.Text += "\n이미지 매칭 성공 ! 클릭준비중";
                        textBox1.Refresh();

                        int x1 = rnd.Next(0, find_img.Width / 10);
                        int y1 = rnd.Next(0, find_img.Height / 10);
                        p.X = maxloc.X + x1;
                        p.Y = maxloc.Y + y1;
                        //Inclick(maxloc.X + x1 , maxloc.Y + y1 , nhwnd);
                        return p;
                    }

                    return p;
                }
            }
            catch {
                
            }
            finally
            {
                rnd = null;
                ScreenMat.Dispose();
                FindMat.Dispose();
                res.Dispose();
            }
            return p;
               
        }




        public bool Inclick(int x, int y, IntPtr nhwnd)
        {
            bool flag = false;
            if( nhwnd != IntPtr.Zero)
            {
                Random r = new Random();
                Task.Delay(r.Next(0, 1000)).Wait();
                IntPtr lparam = new IntPtr(x | (y << 16));
                try
                {
                    if (lparam != IntPtr.Zero)
                    {
                        textBox2.Text = "해당이미지 클릭";
                        textBox2.Refresh();
                        IntPtr nhwnd2 = FindWindowEx((int)nhwnd, 0, "Qt5QWindowIcon", "ScreenBoardClassWindow");

                        PostMessage(nhwnd2, WM_LBUTTONDOWN, 1, lparam);
                        Task.Delay(500).Wait();
                        PostMessage(nhwnd2, WM_LBUTTONUP, 0, lparam);
                        flag = true;
                        nhwnd2 = IntPtr.Zero;
                    }

                }
                catch { }
                finally
                {
                    r = null;
                    lparam = IntPtr.Zero;
                }
            }
            return flag;
        }

 

        public bool buyFlag()
        {
            Bitmap[] thunderImage =
            {
                new Bitmap(@"img\번충안내.PNG"),
                new Bitmap(@"img\번충예.PNG"),
                new Bitmap(@"img\에너지충전.PNG"),
                new Bitmap(@"img\번충예2.PNG"),
                new Bitmap(@"img\번충확인.PNG"),
                new Bitmap(@"img\번충닫기.PNG")
            };
            try
            {
                for (int b = 0; b < thunderImage.Length; b++)
                {
                    Task.Delay(2000).Wait();
                    textBox1.Text = "번개 충전(0 부터 5까지)" + b;
                    textBox1.Refresh();
                    Bitmap screen_img = PrintWindow2();
                    p2 = searchImg(screen_img, thunderImage[b]);
                    if (Inclick(p2.X, p2.Y, nhwnd) != true)
                    {
                        thunderImage[0].Dispose();
                        thunderImage[1].Dispose();
                        thunderImage[2].Dispose();
                        thunderImage[3].Dispose();
                        thunderImage[4].Dispose();
                        thunderImage[5].Dispose();
                        buyFlag();
                    }

                }
            }
            catch { }
            finally {
                thunderImage[0].Dispose();
                thunderImage[1].Dispose();
                thunderImage[2].Dispose();
                thunderImage[3].Dispose();
                thunderImage[4].Dispose();
                thunderImage[5].Dispose();
            }

            return true;
            
        }

        private bool pointclick(int X, int Y,int Xsize, int Ysize, string description)
        {
            System.Drawing.Point p = new System.Drawing.Point();
            Random r = new Random();
            try
            {
                textBox1.Text = description;
                textBox1.Refresh();
                p2.X = X + r.Next(0, Xsize/2 );
                p2.Y = Y + r.Next(0, Ysize/2 );
                bool flag = Inclick(p2.X, p2.Y, nhwnd);
                Task.Delay(1000 + r.Next(0, 500)).Wait();
                return flag;
            }
            catch { }
            finally {
                r = null;
                p = System.Drawing.Point.Empty;
            }
            return false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //Bitmap sell = new Bitmap(@"img\판매.PNG");
            //Bitmap star5 = new Bitmap(@"img\5성희귀.PNG");
            textBox2.Text = "";
            textBox2.Refresh();


            Random r = new Random();
            Bitmap screen_img = PrintWindow2();
            Bitmap start_img = new Bitmap(@"img\전투시작.PNG");
            try
            {
                int i = 0;
                
                textBox1.Text = "승리 패배 검사";
                textBox1.Refresh();
                Bitmap[] winOrLose =
                {
                    new Bitmap(@"img\victory.PNG"),
                    new Bitmap(@"img\전복.PNG")
                };
                try
                {
                    for (; i < winOrLose.Length; i++)
                    {
                        screen_img = PrintWindow2();
                        p2 = searchImg(screen_img, winOrLose[i]);
                        if (Inclick(p2.X, p2.Y, nhwnd))
                        {
                            break;
                        }
                    }
                }
                catch { }
                finally
                {
                    winOrLose[0].Dispose();
                    winOrLose[1].Dispose();
                    winOrLose = null;
                }
                screen_img.Dispose();
                
                if (i == 0)
                {
                    
                    textBox1.Text = "승리";
                    textBox1.Refresh();
                    screen_img = PrintWindow2();
                    pointclick(p2.X, p2.Y, 160, 73, "아무데나 클릭");

                    Bitmap check_rare_5star = new Bitmap(@"img\check_rare_5star.PNG");
                    p2 = searchImg(screen_img, check_rare_5star);
                    if (!p2.IsEmpty && starFlag_check.Checked == true)
                    {
                        Bitmap sell_runes = new Bitmap(@"img\sell_runes.PNG");
                        Bitmap sell_runes_yes = new Bitmap(@"img\sell_runes_yes.PNG");

                        p2 = searchImg(screen_img, sell_runes);
                        if (!p2.IsEmpty)
                        {
                            pointclick(p2.X, p2.Y, sell_runes.Width, sell_runes.Height, "5성룬 판매");
                            screen_img = PrintWindow2();
                            p2 = searchImg(screen_img, sell_runes_yes);
                            if (p2.IsEmpty)
                            {
                                pointclick(p2.X, p2.Y, sell_runes.Width, sell_runes.Height, "5성룬 판매 확인");
                                Task.Delay(1000).Wait();
                            }

                        }
                        sell_runes.Dispose();

                    }
                    else
                    {
                        pointclick(416, 441, 119, 47, "룬획득");
                    }
                    Bitmap getItem_image = new Bitmap(@"img\아이템_확인.PNG");
                    p2 = searchImg(screen_img, getItem_image);
                    if (!p2.IsEmpty)
                    {
                        pointclick(p2.X, p2.Y, getItem_image.Width, getItem_image.Height, "아이템 확인");
                    }
                    Task.Delay(1000).Wait();
                    pointclick(132, 314, 227, 45, "다시하기");
                    screen_img.Dispose();
                    getItem_image.Dispose();
                    check_rare_5star.Dispose();
                    
                }
                else if(i==1)
                {
                    textBox1.Text = "전복";
                    textBox1.Refresh();
                    Bitmap loseImage = new Bitmap(@"img\전복아니오.PNG");
                    screen_img = PrintWindow2();
                    p2 = searchImg(screen_img, loseImage);
                    if (!p2.IsEmpty)
                    {
                        pointclick(p2.X, p2.Y, loseImage.Width, loseImage.Height, "전복 아니오");
                    }
                    loseImage.Dispose();
                    pointclick(132, 314, 227, 45, "다시하기");
                    screen_img.Dispose();
                }
                else
                {

                }
                Task.Delay(1000).Wait();
                Bitmap thunderbuy = new Bitmap(@"img\번충안내.PNG");
                Bitmap replayImage = new Bitmap(@"img\다시하기.PNG");
                screen_img = PrintWindow2();

                p2 = searchImg(screen_img, thunderbuy);
                screen_img.Dispose();
                if (!p2.IsEmpty)
                {
                    buyFlag();
                    screen_img = PrintWindow2();
                        
                    p2 = searchImg(screen_img, replayImage);
                    Inclick(p2.X, p2.Y, nhwnd);
                    pointclick(132, 314, 227, 45, "다시하기");
                }
                
                replayImage.Dispose();
                thunderbuy.Dispose();
                screen_img.Dispose();
                
                screen_img = PrintWindow2();
                p2 = searchImg(screen_img, start_img);
                Task.Delay(600).Wait();
                if (!p2.IsEmpty)
                {
                    pointclick(p2.X, p2.Y, start_img.Width, start_img.Height, "전투시작");
                }
                textBox1.Text = "한바퀴 끝 1분 후 다음 반복 시작";
                textBox1.Refresh();
                start_img.Dispose();

            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message.ToString());
            }
            finally
            {

                screen_img.Dispose();
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
