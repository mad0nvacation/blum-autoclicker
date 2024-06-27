using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using OpenCvSharp;
using WindowsInput;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Point = System.Drawing.Point; // Resolve namespace conflict by aliasing

namespace WinFormsApp4
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetCursorPos(out Point lpPoint);

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        const int MOUSEEVENTF_LEFTDOWN = 0x02;
        const int MOUSEEVENTF_LEFTUP = 0x04;

        string templatePath = ""; // ???? ? ?????????? ??????????? ??? ??????

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(templatePath))
            {
                MessageBox.Show("???????? ??????????? ??? ??????.");
                return;
            }

            // Start the main logic on a separate thread to avoid freezing the UI
            Thread mainThread = new Thread(MainLogic);
            mainThread.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "???????? ??????????? ??? ??????";
            openFileDialog.Filter = "??????????? (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|??? ????? (*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                templatePath = openFileDialog.FileName;

               
            }
        }

        private void MainLogic()
        {
            string screenshotPath = Path.GetTempPath() + "screenshot.png";
            Bitmap screenshot;

            // Colors to find and click
            string[] hexColors = {
                "#c9e100",
                "#bae70e",
                "#abff61",
                "#87ff27"
            };
            Color[] colorsToFind = new Color[hexColors.Length];
            for (int i = 0; i < hexColors.Length; i++)
            {
                colorsToFind[i] = HexToColor(hexColors[i]);
            }

            Stopwatch stopwatch = new Stopwatch();
            DateTime lastTemplateSearchTime = DateTime.MinValue;

            while (true)
            {
                // Capture left half of the screen for color search
                screenshot = CaptureLeftHalfScreen();

                // Find colors and click
                foreach (Color color in colorsToFind)
                {
                    stopwatch.Restart();
                    Point? point = FindColor(screenshot, color);
                    stopwatch.Stop();

                    if (point.HasValue)
                    {
                        // Click at the found point
                        Click(point.Value);
                    }
                }

                // Dispose screenshot after processing for color search
                screenshot.Dispose();

                // Check if 10 seconds have passed since the last template search
                if ((DateTime.Now - lastTemplateSearchTime).TotalSeconds >= 10)
                {
                    // Capture the whole screen for template matching
                    screenshot = CaptureScreen();

                    // Save the screenshot for template matching
                    screenshot.Save(screenshotPath, ImageFormat.Png);

                    // Image template matching
                    Mat template = Cv2.ImRead(templatePath, ImreadModes.Color);
                    Mat screen = Cv2.ImRead(screenshotPath, ImreadModes.Color);

                    // Perform template matching
                    Mat result = new Mat();
                    Cv2.MatchTemplate(screen, template, result, TemplateMatchModes.CCoeffNormed);
                    Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

                    // Define threshold for template matching
                    double threshold = 0.8;

                    if (maxVal >= threshold)
                    {
                        // Coordinates for mouse click
                        int clickX = maxLoc.X + template.Width / 2;
                        int clickY = maxLoc.Y + template.Height / 2;

                        // Simulate mouse click
                        var sim = new InputSimulator();
                        sim.Mouse.MoveMouseToPositionOnVirtualDesktop(clickX * 65535 / Screen.PrimaryScreen.Bounds.Width, clickY * 65535 / Screen.PrimaryScreen.Bounds.Height);
                        sim.Mouse.LeftButtonClick();
                    }

                    // Dispose screenshot after processing for template matching
                    screenshot.Dispose();

                    // Update the last template search time
                    lastTemplateSearchTime = DateTime.Now;
                }

                // Small delay to avoid CPU overload
                Thread.Sleep(1);
            }
        }

        static void Click(Point point)
        {
            SetCursorPos(point.X, point.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, point.X, point.Y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, point.X, point.Y, 0, 0);
        }

        static Bitmap CaptureScreen()
        {
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
            }
            return bitmap;
        }

        static Bitmap CaptureLeftHalfScreen()
        {
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            bounds.Width /= 2;
            Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
            }
            return bitmap;
        }

        static Point? FindColor(Bitmap bitmap, Color color)
        {
            BitmapData bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, bitmap.PixelFormat);
            int bytesPerPixel = Image.GetPixelFormatSize(bitmap.PixelFormat) / 8;
            int heightInPixels = bitmapData.Height;
            int widthInBytes = bitmapData.Width * bytesPerPixel;
            Point? result = null;

            unsafe
            {
                byte* ptrFirstPixel = (byte*)bitmapData.Scan0;

                for (int y = 0; y < heightInPixels && result == null; y++)
                {
                    byte* currentLine = ptrFirstPixel + (y * bitmapData.Stride);
                    for (int x = 0; x < widthInBytes; x += bytesPerPixel)
                    {
                        int b = currentLine[x];
                        int g = currentLine[x + 1];
                        int r = currentLine[x + 2];
                        Color pixelColor = Color.FromArgb(r, g, b);

                        if (IsColorMatch(pixelColor, color))
                        {
                            result = new Point(x / bytesPerPixel, y);
                            break;
                        }
                    }
                }
            }

            bitmap.UnlockBits(bitmapData);
            return result;
        }

        static bool IsColorMatch(Color c1, Color c2)
        {
            return Math.Abs(c1.R - c2.R) < 20 && Math.Abs(c1.G - c2.G) < 20 && Math.Abs(c1.B - c2.B) < 20;
        }

        static Color HexToColor(string hexColor)
        {
            hexColor = hexColor.Replace("#", "");
            int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
            int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
            int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
            return Color.FromArgb(r, g, b);
        }
    }
}
