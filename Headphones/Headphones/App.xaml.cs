using System.Windows;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Vannatech.CoreAudio.Constants;
using Vannatech.CoreAudio.Enumerations;
using Vannatech.CoreAudio.Externals;
using Vannatech.CoreAudio.Interfaces;
using Vannatech.CoreAudio.Structures;

namespace Headphones
{
    /// <summary>
    /// Simple application. Check the XAML for comments.
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private TaskbarIcon notifyIcon;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static IAudioEndpointVolume audioEndpointVolume;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            //create the notifyicon
            notifyIcon = (TaskbarIcon)FindResource("NotifyIcon");

            _hookID = SetHook(_proc);

            var clsid = new Guid(ComCLSIDs.MMDeviceEnumeratorCLSID);
            var deviceEnumeratorType = Type.GetTypeFromCLSID(clsid);
            var deviceEnumerator = (IMMDeviceEnumerator)
                Activator.CreateInstance(deviceEnumeratorType);

            IMMDevice device;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender,
                ERole.eMultimedia, out device);

            var iid = new Guid(ComIIDs.IAudioEndpointVolumeIID);
            var context = (uint)CLSCTX.CLSCTX_INPROC_SERVER;
            object objInterface;
            device.Activate(iid, context, IntPtr.Zero, out objInterface);
            audioEndpointVolume = (IAudioEndpointVolume)objInterface;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            //the icon would clean up automatically, but this is cleaner
            notifyIcon.Dispose();
            UnhookWindowsHookEx(_hookID);
            base.OnExit(e);
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }
        
        private delegate IntPtr LowLevelKeyboardProc(
            int nCode, IntPtr wParam, IntPtr lParam);

        private static float previousLevel = 0;

        private static IntPtr HookCallback(
            int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                var mods = Keys.Control | Keys.Shift;

                if ((Control.ModifierKeys & mods) == mods) {

                    float level;
                    audioEndpointVolume.GetMasterVolumeLevelScalar(out level);

                    switch ((Keys)vkCode) {
                        case Keys.F10:
                            if (previousLevel == 0)
                            {
                                SetVolume(0);
                                previousLevel = level;
                            }
                            else
                            {
                                SetVolume(previousLevel);
                                previousLevel = 0;
                            }
                            break;

                        case Keys.F11:
                            SetVolume(level - 0.01f);
                            previousLevel = 0;
                            break;

                        case Keys.F12:
                            SetVolume(level + 0.01f);
                            previousLevel = 0;
                            break;
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private static void SetVolume(float level)
        {
            level = Math.Min(Math.Max(level, 0), 100);
            var context = Guid.NewGuid();
            audioEndpointVolume.SetMasterVolumeLevelScalar(level, context);
        }
    }
}
