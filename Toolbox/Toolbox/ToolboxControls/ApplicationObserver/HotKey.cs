using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ApplicationObserver
{
    /// <summary>
    /// global hotkey
    /// taken from http://www.mycsharp.de/wbb2/thread.php?threadid=65760
    /// </summary>
    public class Hotkey : IDisposable
    {
        #region Fields

        private Keys _keys;
        private int _id;

        #endregion

        #region Events
        
        /// <summary>
        /// Rises when the hotkey is pressed
        /// </summary>
        public event EventHandler HotkeyPressed;

        #endregion

        #region Construction

        public Hotkey()
        { }

        #endregion

        #region Properties

        public object Tag { get; set; }

        /// <summary>
        /// The keycombination
        /// </summary>
        public Keys Keys
        {
            get { return _keys; }
        }
        
        #endregion

        #region Methods

        /// <summary>
        /// Registers the hotkey. You have to keep a reference to the returned object.
        /// </summary>
        /// <param name="keys"></param>
        /// <returns></returns>
        public static Hotkey Register(Keys keys)
        {
            Hotkey ret = new Hotkey();
            ret._keys = keys;
            wnd.Register(ret);
            return ret;
        }

        /// <summary>
        /// Calls Dispose: Unregisters the hotkey
        /// </summary>
        /// <param name="h">The Hotkey</param>
        public static void UnRegister(Hotkey h)
        { 
            h.Dispose();
        }
        
        #endregion

        #region IDisposable Member

        bool _disposed = false;

        /// <summary>
        /// Unregisters the Hotkey
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            wnd.UnRegister(this);
        }

        #endregion

        #region Private class wnd

        private class wnd : Control
        {
            public wnd()
            {
                Visible = false;
            }

            public static wnd Default
            {
                get
                {
                    if (def == null)
                    {
                        def = new wnd();
                        def.CreateHandle();
                    }
                    return def;
                }
            }

            List<IntPtr> hotkeys = new List<IntPtr>();
            int getNewId(IntPtr item)
            {
                int i = 0;
                foreach (IntPtr r in hotkeys)
                {
                    if ((long)r == 0)
                    {
                        hotkeys[i] = item;
                        return i;
                    }
                    i++;
                }
                hotkeys.Add(item);
                return i;
            }

            IntPtr getObject(int id)
            {
                return hotkeys[id];
            }

            void removeId(int id)
            {
                hotkeys[id] = (IntPtr)0;
            }

            static wnd def;
            public short LOWORD(int l)
            { return ((short)(l & 0xffff)); }

            public short HIWORD(int l)
            { return ((short)(l >> 16)); }

            public const int MOD_ALT = 0x1;
            public const int MOD_CONTROL = 0x2;
            public const int MOD_SHIFT = 0x4;
            public const int MOD_WIN = 0x8;
            public const int WM_HOTKEY = 0x312;

            [DllImport("user32.dll")]
            private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

            [DllImport("user32.dll")]
            private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    Hotkey h = (Hotkey)GCHandle.FromIntPtr(getObject((int)m.WParam)).Target;
                    if (h.HotkeyPressed != null)
                    { h.HotkeyPressed(h, null); }
                }
                else
                    base.WndProc(ref m);
            }

            internal static void Register(Hotkey h)
            {
                h._id = Default.getNewId(GCHandle.ToIntPtr(GCHandle.Alloc(h, GCHandleType.WeakTrackResurrection)));
                int modifiers = 0;
                if ((h._keys & Keys.Alt) == Keys.Alt)
                    modifiers = modifiers | MOD_ALT;
                if ((h._keys & Keys.Control) == Keys.Control)
                    modifiers = modifiers | MOD_CONTROL;
                if ((h._keys & Keys.Shift) == Keys.Shift)
                    modifiers = modifiers | MOD_SHIFT;
                Keys k = h._keys & ~Keys.Control & ~Keys.Shift & ~Keys.Alt;
                RegisterHotKey((IntPtr)Default.Handle, h._id, modifiers, (int)k);
            }

            internal static void UnRegister(Hotkey h)
            {
                try
                {
                    UnregisterHotKey(Default.Handle, h._id);
                }
                catch { }
                GCHandle.FromIntPtr(Default.getObject(h._id)).Free();
                Default.removeId(h._id);
            }
        }

        #endregion
    } 
}
