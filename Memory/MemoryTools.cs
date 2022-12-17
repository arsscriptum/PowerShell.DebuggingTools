using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MemoryTools.Win32
{
  
    public static class Memory
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool GetProcessWorkingSetSizeEx( IntPtr hProcess, ref int min, ref int max , ref int flags );
    }
    public static class UserInput
    {  
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public int dwTime;
        }
        public static DateTime LastInput
        {
            get
            {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }
        public static TimeSpan IdleTime
        {
            get
            {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }
        public static int LastInputTicks
        {
            get
            {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}