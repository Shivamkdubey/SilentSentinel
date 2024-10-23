using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoMover
{
    public class MoveMouse : Form
    {       
        #region Constants
        private int angle = 0;
        private const int radius = 100;
        private const int centerX = 1200;
        private const int centerY = 250;
        private const int timerInterval = 10 * 1000;        // 10 secs in ms
        private const int idleTimerInterval = 60 * 1000;    // 60 secs in ms
        private const int idleThresholdInterval = 4 * 60;   // 4 minutes in sec
        private const uint MOUSEEVENT_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENT_LEFTUP = 0x0004;
        private int Angle
        {
            get => angle;
            set
            {
                angle = value;
                if (angle >= 360)
                {
                    angle -= 360;
                }
            }
        }
        #endregion

        #region Member Fields
        private Button toggleButton;
        private Timer idleTimer;
        private Timer timer;
        private bool isRunning = false;
        private bool isAutomatedClick = false;

        #endregion

        #region Global Mouse Hook
        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_MOUSE_LL = 14;
        private const int WM_MOUSEMOVE = 0x0200;
        private static DateTime lastActivityTime = DateTime.Now;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelMouseProc mouseProc;
        private IntPtr hookID = IntPtr.Zero;

        private void SetMouseHook()
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                mouseProc = HookCallback;
                hookID = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private void UnhookMouse()
        {
            UnhookWindowsHookEx(hookID);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEMOVE)
            {
                lastActivityTime = DateTime.Now;
            }

            return CallNextHookEx(hookID, nCode, wParam, lParam);
        }
        #endregion

        #region Global Keyboard Hook
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc keyboardProc;
        private IntPtr keyboardHookID = IntPtr.Zero;

        private void SetKeyboardHook()
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                keyboardProc = KeyboardHookCallback;
                keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private void UnhookKeyboard()
        {
            UnhookWindowsHookEx(keyboardHookID);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                lastActivityTime = DateTime.Now;
            }

            return CallNextHookEx(keyboardHookID, nCode, wParam, lParam);
        }
        #endregion

        #region Maximize/Minimize Windows
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)]  
        static extern IntPtr SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);

        const int WM_COMMAND = 0x111;
        const int MIN_ALL = 419;
        const int MIN_ALL_UNDO = 416;

        private void MinimizeAllWindows()
        {       
            IntPtr lHwnd = FindWindow("Shell_TrayWnd", null);
            SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL, IntPtr.Zero);
            System.Threading.Thread.Sleep(500);
        }
        #endregion

        public MoveMouse()
        {
            InitializeComponents();
            PlaceComponents();
            UpdateFormColor();
            SetMouseHook();
            SetKeyboardHook();
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

            Start();
        }

        private async void Start()
        {
            await Task.Delay(100);
            ToggleRunningState(null, null);
        }

        #region Create App Components
        private void InitializeComponents()
        {
            toggleButton = new Button();
            toggleButton.Text = "Stop";
            toggleButton.Click += ToggleRunningState;
            CustomizeButton(toggleButton);

            timer = new Timer
            {
                Interval = timerInterval
            };
            idleTimer = new Timer 
            {
                Interval = idleTimerInterval
            };
            timer.Tick += Timer_Tick;
            idleTimer.Tick += IdleTimer_Tick;

            Controls.Add(toggleButton);

            Text = "Silent Sentinel";
            Size = new Size(600, 500);

            void CustomizeButton(Button button)
            {
                button.Width = 150;
                button.Height = 50;
                toggleButton.Font = new Font(toggleButton.Font.FontFamily, 12, FontStyle.Bold);
                button.FlatStyle = FlatStyle.Standard;
                button.BackColor = Color.Red;
                button.ForeColor = Color.White;
            }
        }

        private void PlaceComponents()
        {
           
            toggleButton.Left = (ClientSize.Width - toggleButton.Width) / 2;
            toggleButton.Top = (ClientSize.Height - toggleButton.Height) / 5;
        }
        #endregion
                
        #region Business Logic
        private void Timer_Tick(object sender, EventArgs e)
        {
            ClickMouse();
            Angle += 36;
        }

        private void IdleTimer_Tick(object sender, EventArgs e)
        {
            //print($"{DateTime.Now.Subtract(lastActivityTime).TotalSeconds} running = {isRunning}");
            if (!isRunning && DateTime.Now.Subtract(lastActivityTime).TotalSeconds >= idleThresholdInterval)
            {
                RestartAfterIdleState();
            }
        }
                
        private void RestartAfterIdleState()
        {
            MinimizeAllWindows();
            WakeUpScreen();     
            Activate();
            System.Threading.Thread.Sleep(500);
            MinimizeAllWindows();
            ToggleRunningState(null, null);     
        }

        private void ToggleRunningState(object sender, EventArgs e)
        {
            isRunning = !isRunning;
            if (isRunning)
            {
                timer.Start();
                toggleButton.Text = "Stop";
                toggleButton.BackColor = Color.Red;
                ClickMouse();

                idleTimer.Stop();
            }
            else
            {
                timer.Stop();
                toggleButton.Text = "Start";
                toggleButton.BackColor = Color.Green;

                idleTimer.Start();
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            const int WM_ACTIVATEAPP = 0x001C;

            switch (m.Msg)
            {
                case WM_ACTIVATEAPP:
                    // Check if the app is being deactivated (loses focus)
                    bool appDeactivated = (int)m.WParam == 0;

                    if (appDeactivated && !isAutomatedClick && isRunning)
                    {
                        ToggleRunningState(null, null);
                    }
                    break;
            }
        }
        #endregion

        #region Session/Power Logic
        private void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                lastActivityTime = DateTime.Now;
                idleTimer.Stop();
            }
            //print($"{e.Mode.ToString().ToUpper()} {lastActivityTime}");
        }
        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            lastActivityTime = DateTime.Now;
            if (e.Reason == SessionSwitchReason.SessionLock || e.Reason == SessionSwitchReason.SessionLogoff)
            {
                idleTimer.Stop();
            }
            if (e.Reason == SessionSwitchReason.SessionUnlock || e.Reason == SessionSwitchReason.SessionLogon)
            {
                idleTimer.Start();
            }
            //print($"{e.Reason.ToString().ToUpper()} {lastActivityTime}");
        }
        #endregion

        #region Utility
        private void ClickMouse()
        {
            isAutomatedClick = true;

            int newX = centerX + (int)(radius * Math.Cos(Angle * Math.PI / 180));
            int newY = centerY + (int)(radius * Math.Sin(Angle * Math.PI / 180));

            Cursor.Position = new Point(newX, newY);
            mouse_event(MOUSEEVENT_LEFTDOWN | MOUSEEVENT_LEFTUP, 0, 0, 0, 0);

            isAutomatedClick = false;

            Activate();
        }


        private async void BringFormToTop()
        {
            TopMost = true;
            Activate();
            await Task.Delay(3000);
            TopMost = false; // Restore normal behavior after bringing the form to the top
        }

        private static void WakeUpScreen()
        {
            // Wakes Screen if turned off, not asleep
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
        }
        #endregion

        #region App Theme
        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
            {
                UpdateFormColor();
            }
        }

        private void UpdateFormColor()
        {
            if (IsDarkMode())
            {
                BackColor = Color.FromArgb(31, 31, 31);
                ForeColor = Color.White;
            }
            else
            {
                BackColor = SystemColors.Control;
                ForeColor = SystemColors.ControlText;
            }
        }

        private bool IsDarkMode()
        {
            const string keyName = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    if (value != null && value is int intValue)
                    {
                        return intValue == 0;
                    }
                }
                return true;
            }
        }
        #endregion
                
        #region Dispose Method
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                UnhookMouse();
                UnhookKeyboard();
            }
            base.Dispose(disposing);
        }
        #endregion

        private void print(object message)
        {
            Debug.WriteLine(message);
        }
    }
}
