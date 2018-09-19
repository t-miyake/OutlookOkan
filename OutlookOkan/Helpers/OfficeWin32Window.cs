using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookOkan.Helpers
{
    public class OfficeWin32Window : IWin32Window
    {
        [DllImport("user32", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public IntPtr Handle { get; }

        public OfficeWin32Window(object windowObject)
        {
            Handle = FindWindow("rctrl_renwnd32\0", windowObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, windowObject, null).ToString());
        }
    }
}