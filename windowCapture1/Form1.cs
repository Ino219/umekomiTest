using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Shell32;
using SHDocVw;
using System.IO;


namespace windowCapture1
{
    
    public partial class Form1 : Form
    {
        private const int SRCCOPY = 13369376;
        private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

        static async void delayTask()
        {

            await Task.Delay(1000);

        }

        Button button1;
        PictureBox pictureBox;


        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }


        [DllImport("user32.dll")]
        static extern IntPtr GetWindowRect(IntPtr hWnd, out Rect rect);


        [DllImport("user32.dll")]
        extern static IntPtr GetForegroundWindow();


        [DllImport("dwmapi.dll")]
        extern static int DwmGetWindowAttribute(IntPtr hWnd, int dwAttribute, out Rect rect, int cbAttribute);


        [DllImport("user32.dll")]
        static extern IntPtr GetWindowDC(IntPtr hWnd);


        [DllImport("gdi32.dll")]
        static extern int BitBlt(IntPtr hDestDC, int x, int y, int nWidth, int nHeight, IntPtr hSrcDC, int xSrc, int ySrc, int dwRop);


        [DllImport("user32.dll")]
        static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hdc);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("shell32.dll", SetLastError = true)]
        public static extern int SHOpenFolderAndSelectItems(IntPtr pidlFolder, uint cidl, [In, MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, uint dwFlags);

        [DllImport("shell32.dll")]
        public static extern void SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr bindingContext, [Out()] out IntPtr pidl, uint sfgaoIn, [Out()] out uint psfgaoOut);

        public Form1()
        {
            InitializeComponent();
            this.AutoScroll = true;
            this.Load += Form1_Load;

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            button1 = new Button();
            button1.Location = new Point(10, 10);
            button1.Text = "アクティブウィンドウ";
            button1.AutoSize = true;
            button1.Click += Button1_Click;


            pictureBox = new PictureBox();
            pictureBox.Location = new Point(10, 40);
            pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;


            this.Controls.Add(button1);
            this.Controls.Add(pictureBox);
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            Rect bounds, rect;
            // アクティブウィンドウを取得
            IntPtr hWnd = GetForegroundWindow();
            IntPtr winDC = GetWindowDC(hWnd);
            DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, out bounds, Marshal.SizeOf(typeof(Rect)));
            GetWindowRect(hWnd, out rect);


            var offsetX = bounds.left - rect.left;
            var offsetY = bounds.top - rect.top;


            Bitmap bitmap = new Bitmap(bounds.right - bounds.left, bounds.bottom - bounds.top);
            Graphics graphics = Graphics.FromImage(bitmap);
            IntPtr hDC = graphics.GetHdc();


            BitBlt(hDC, 0, 0, bitmap.Width, bitmap.Height, winDC, offsetX, offsetY, SRCCOPY);
            graphics.ReleaseHdc(hDC);
            graphics.Dispose();
            ReleaseDC(hWnd, winDC);


            // 表示
            pictureBox.Image = bitmap;
        }
        //現在のアクティブウィンドウを取得して、画像として保存する
        public static string GetGazouPath()
        {
            //アクティブウィンドウの取得
            Rect bounds, rect;
            // アクティブウィンドウのハンドルを取得
            IntPtr hWnd = GetForegroundWindow();
            //ウィンドウハンドルからタイトルバーやメニューなどのデバイスコンテクストを取得
            IntPtr winDC = GetWindowDC(hWnd);
            DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, out bounds, Marshal.SizeOf(typeof(Rect)));
            GetWindowRect(hWnd, out rect);

            //画面調整用のオフセットを設定
            var offsetX = bounds.left - rect.left;
            var offsetY = bounds.top - rect.top;

            //描画用のbitmapデータを設定
            Bitmap bitmap = new Bitmap(bounds.right - bounds.left, bounds.bottom - bounds.top);
            Graphics graphics = Graphics.FromImage(bitmap);
            //Graphicsに設定したbitmapデータのハンドルを取得
            IntPtr hDC = graphics.GetHdc();
            
            BitBlt(hDC, 0, 0, bitmap.Width, bitmap.Height, winDC, offsetX, offsetY, SRCCOPY);
            //リソースの解放
            graphics.ReleaseHdc(hDC);
            graphics.Dispose();
            ReleaseDC(hWnd, winDC);
            //保存先のパスを指定
            string savePath = "C:\\Users\\chach\\Desktop\\test.png";
            //保存先のパスにbitmapを保存
            bitmap.Save(savePath);
            //保存先のパスを返す
            return savePath;
        }
        public static List<string> getOpenFolderList()
        {
            List<string> ofList = new List<string>();

            //COMのShellクラス作成
            Shell shell = new Shell();
            //IEとエクスプローラの一覧を取得
            ShellWindows win = shell.Windows();
            //列挙
            foreach (IWebBrowser2 web in win)
            {
                //エクスプローラのみ(IEを除外)
                if (Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                {
                    string localPath= new Uri(web.LocationURL).LocalPath;
                    //リストに追加
                    //ofList.Add(web.LocationURL + " : " + web.LocationName);
                    ofList.Add(localPath);
                }
            }
            return ofList;
        }
        public static IntPtr FindWindowByName(string windowName)
        {
            IntPtr hwnd_ = FindWindow(null,windowName);
            return hwnd_;
        }

        public static void SetForeGroundByIntPtr(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
        }

        public static void setForeWindowSelectFile(string filename)
        {
            IntPtr pidl;
            uint flags;
            SHParseDisplayName(filename, IntPtr.Zero, out pidl, 0, out flags);
            MessageBox.Show(pidl.ToString());
            MessageBox.Show(flags.ToString());

            SHOpenFolderAndSelectItems(pidl, 0, null, 0);
        }
        
        public static async Task AsyncFunction()
        {
            
            // 非同期関数を4秒停止
            await Task.Delay(4000);
        }

    }
}
