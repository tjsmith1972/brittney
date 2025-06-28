using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Windows.Forms;

namespace VoiceScreenCapture
{
    public class ScreenCaptureApp
    {
        private SpeechRecognitionEngine _recognizer;

        public void StartListening()
        {
            _recognizer = new SpeechRecognitionEngine();
            
            // Create the grammar with the activation phrase
            var grammarBuilder = new GrammarBuilder("Hey Brittney shoot");
            var grammar = new Grammar(grammarBuilder);
            
            _recognizer.LoadGrammar(grammar);
            _recognizer.SpeechRecognized += Recognizer_SpeechRecognized;
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
            
            Console.WriteLine("Listening for 'Hey Brittney shoot'...");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private void Recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "Hey Brittney shoot")
            {
                Console.WriteLine("Activation phrase detected. Preparing to capture screen...");
                
                // Show form to let user select window
                using (var form = new WindowSelectionForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        CaptureWindow(form.SelectedWindowHandle);
                    }
                }
            }
        }

        private void CaptureWindow(IntPtr windowHandle)
        {
            // Get the window rectangle
            GetWindowRect(windowHandle, out RECT windowRect);

            int width = windowRect.Right - windowRect.Left;
            int height = windowRect.Bottom - windowRect.Top;

            // Create bitmap to hold the screenshot
            using (Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bmp))
                {
                    graphics.CopyFromScreen(windowRect.Left, windowRect.Top, 0, 0, new Size(width, height));
                }

                // Copy to clipboard
                Clipboard.SetImage(bmp);
                Console.WriteLine("Screenshot captured and copied to clipboard!");
            }
        }

        // Window selection form
        private class WindowSelectionForm : Form
        {
            public IntPtr SelectedWindowHandle { get; private set; } = IntPtr.Zero;
            private Timer _timer;
            private IntPtr _hoveredWindow;

            public WindowSelectionForm()
            {
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.TopMost = true;
                this.BackColor = Color.FromArgb(100, Color.Gray);
                this.TransparencyKey = Color.Magenta;
                this.Opacity = 0.5;
                this.DoubleBuffered = true;

                _timer = new Timer { Interval = 50 };
                _timer.Tick += Timer_Tick;
                _timer.Start();

                this.MouseClick += WindowSelectionForm_MouseClick;
                this.KeyDown += WindowSelectionForm_KeyDown;
            }

            private void WindowSelectionForm_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            }

            private void Timer_Tick(object sender, EventArgs e)
            {
                var cursorPos = Cursor.Position;
                _hoveredWindow = WindowFromPoint(cursorPos);
                this.Invalidate();
            }

            private void WindowSelectionForm_MouseClick(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left && _hoveredWindow != IntPtr.Zero)
                {
                    SelectedWindowHandle = _hoveredWindow;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                if (_hoveredWindow != IntPtr.Zero)
                {
                    GetWindowRect(_hoveredWindow, out RECT rect);
                    Rectangle windowRect = new Rectangle(
                        rect.Left, 
                        rect.Top, 
                        rect.Right - rect.Left, 
                        rect.Bottom - rect.Top);

                    // Convert screen coordinates to form coordinates
                    Point screenOffset = this.PointToScreen(Point.Empty);
                    windowRect.X -= screenOffset.X;
                    windowRect.Y -= screenOffset.Y;

                    using (Pen pen = new Pen(Color.Red, 3))
                    {
                        e.Graphics.DrawRectangle(pen, windowRect);
                    }
                }

                using (Font font = new Font("Arial", 12, FontStyle.Bold))
                using (SolidBrush brush = new SolidBrush(Color.White))
                {
                    string instructions = "Click on a window to capture it (ESC to cancel)";
                    SizeF textSize = e.Graphics.MeasureString(instructions, font);
                    e.Graphics.DrawString(
                        instructions, 
                        font, 
                        brush, 
                        (this.Width - textSize.Width) / 2, 
                        20);
                }
            }
        }

        // DLL imports for window handling
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point pt);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right; public int Bottom;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            ScreenCaptureApp app = new ScreenCaptureApp();
            app.StartListening();
        }
    }
}
