using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    /// <summary>
    /// Native Window Message to any instance with a valid window handle
    /// </summary>
    public enum WndMessage : int
    {
        /// <summary>
        /// 
        /// </summary>
        WM_NULL = 0x0000,
        /// <summary>
        /// 
        /// </summary>
        WM_CREATE = 0x0001,
        /// <summary>
        /// 
        /// </summary>
        WM_DESTROY = 0x0002,
        /// <summary>
        /// 
        /// </summary>
        WM_MOVE = 0x0003,
        /// <summary>
        /// 
        /// </summary>
        WM_SIZE = 0x0005,
        /// <summary>
        /// 
        /// </summary>
        WM_ACTIVATE = 0x0006,
        /// <summary>
        /// 
        /// </summary>
        WM_SETFOCUS = 0x0007,
        /// <summary>
        /// 
        /// </summary>
        WM_KILLFOCUS = 0x0008,
        /// <summary>
        /// 
        /// </summary>
        WM_ENABLE = 0x000A,
        /// <summary>
        /// 
        /// </summary>
        WM_SETREDRAW = 0x000B,
        /// <summary>
        /// 
        /// </summary>
        WM_SETTEXT = 0x000C,
        /// <summary>
        /// 
        /// </summary>
        WM_GETTEXT = 0x000D,
        /// <summary>
        /// 
        /// </summary>
        WM_GETTEXTLENGTH = 0x000E,
        /// <summary>
        /// 
        /// </summary>
        WM_PAINT = 0x000F,
        /// <summary>
        /// 
        /// </summary>
        WM_CLOSE = 0x0010,
        /// <summary>
        /// 
        /// </summary>
        WM_QUERYENDSESSION = 0x0011,
        /// <summary>
        /// 
        /// </summary>
        WM_QUIT = 0x0012,
        /// <summary>
        /// 
        /// </summary>
        WM_QUERYOPEN = 0x0013,
        /// <summary>
        /// 
        /// </summary>
        WM_ERASEBKGND = 0x0014,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSCOLORCHANGE = 0x0015,
        /// <summary>
        /// 
        /// </summary>
        WM_ENDSESSION = 0x0016,
        /// <summary>
        /// 
        /// </summary>
        WM_SHOWWINDOW = 0x0018,
        /// <summary>
        /// 
        /// </summary>
        WM_WININICHANGE = 0x001A,
        /// <summary>
        /// 
        /// </summary>
        WM_SETTINGCHANGE = 0x001A,
        /// <summary>
        /// 
        /// </summary>
        WM_DEVMODECHANGE = 0x001B,
        /// <summary>
        ///
        /// </summary>
        WM_ACTIVATEAPP = 0x001C,
        /// <summary>
        /// 
        /// </summary>
        WM_FONTCHANGE = 0x001D,
        /// <summary>
        /// 
        /// </summary>
        WM_TIMECHANGE = 0x001E,
        /// <summary>
        /// 
        /// </summary>
        WM_CANCELMODE = 0x001F,
        /// <summary>
        /// 
        /// </summary>
        WM_SETCURSOR = 0x0020,
        /// <summary>
        /// 
        /// </summary>
        WM_MOUSEACTIVATE = 0x0021,
        /// <summary>
        /// 
        /// </summary>
        WM_CHILDACTIVATE = 0x0022,
        /// <summary>
        /// 
        /// </summary>
        WM_QUEUESYNC = 0x0023,
        /// <summary>
        /// 
        /// </summary>
        WM_GETMINMAXINFO = 0x0024,
        /// <summary>
        /// 
        /// </summary>
        WM_PAINTICON = 0x0026,
        /// <summary>
        /// 
        /// </summary>
        WM_ICONERASEBKGND = 0x0027,
        /// <summary>
        /// 
        /// </summary>
        WM_NEXTDLGCTL = 0x0028,
        /// <summary>
        /// 
        /// </summary>
        WM_SPOOLERSTATUS = 0x002A,
        /// <summary>
        /// 
        /// </summary>
        WM_DRAWITEM = 0x002B,
        /// <summary>
        /// 
        /// </summary>
        WM_MEASUREITEM = 0x002C,
        /// <summary>
        /// 
        /// </summary>
        WM_DELETEITEM = 0x002D,
        /// <summary>
        /// 
        /// </summary>
        WM_VKEYTOITEM = 0x002E,
        /// <summary>
        /// 
        /// </summary>
        WM_CHARTOITEM = 0x002F,
        /// <summary>
        /// 
        /// </summary>
        WM_SETFONT = 0x0030,
        /// <summary>
        /// 
        /// </summary>
        WM_GETFONT = 0x0031,
        /// <summary>
        /// 
        /// </summary>
        WM_SETHOTKEY = 0x0032,
        /// <summary>
        /// 
        /// </summary>
        WM_GETHOTKEY = 0x0033,
        /// <summary>
        /// 
        /// </summary>
        WM_QUERYDRAGICON = 0x0037,
        /// <summary>
        /// 
        /// </summary>
        WM_COMPAREITEM = 0x0039,
        /// <summary>
        /// 
        /// </summary>
        WM_COMPACTING = 0x0041,
        /// <summary>
        /// no longer suported
        /// </summary>
        WM_COMMNOTIFY = 0x0044,  /* no longer suported */
        /// <summary>
        /// 
        /// </summary>
        WM_WINDOWPOSCHANGING = 0x0046,
        /// <summary>
        /// 
        /// </summary>
        WM_WINDOWPOSCHANGED = 0x0047,
        /// <summary>
        /// 
        /// </summary>
        WM_POWER = 0x0048,
        /// <summary>
        /// 
        /// </summary>
        WM_COPYDATA = 0x004A,
        /// <summary>
        /// 
        /// </summary>
        WM_CANCELJOURNAL = 0x004B,
        /// <summary>
        /// 
        /// </summary>
        WM_NOTIFY = 0x004E,
        /// <summary>
        /// 
        /// </summary>
        WM_INPUTLANGCHANGEREQUEST = 0x0050,
        /// <summary>
        /// 
        /// </summary>
        WM_INPUTLANGCHANGE = 0x0051,
        /// <summary>
        /// 
        /// </summary>
        WM_TCARD = 0x0052,
        /// <summary>
        /// 
        /// </summary>
        WM_HELP = 0x0053,
        /// <summary>
        /// 
        /// </summary>
        WM_USERCHANGED = 0x0054,
        /// <summary>
        /// 
        /// </summary>
        WM_NOTIFYFORMAT = 0x0055,
        /// <summary>
        /// 
        /// </summary>
        WM_CONTEXTMENU = 0x007B,
        /// <summary>
        /// 
        /// </summary>
        WM_STYLECHANGING = 0x007C,
        /// <summary>
        /// 
        /// </summary>
        WM_STYLECHANGED = 0x007D,
        /// <summary>
        /// 
        /// </summary>
        WM_DISPLAYCHANGE = 0x007E,
        /// <summary>
        /// 
        /// </summary>
        WM_GETICON = 0x007F,
        /// <summary>
        /// 
        /// </summary>
        WM_SETICON = 0x0080,
        /// <summary>
        /// 
        /// </summary>
        WM_NCCREATE = 0x0081,
        /// <summary>
        /// 
        /// </summary>
        WM_NCDESTROY = 0x0082,
        /// <summary>
        /// 
        /// </summary>
        WM_NCCALCSIZE = 0x0083,
        /// <summary>
        /// 
        /// </summary>
        WM_NCHITTEST = 0x0084,
        /// <summary>
        /// 
        /// </summary>
        WM_NCPAINT = 0x0085,
        /// <summary>
        /// 
        /// </summary>
        WM_NCACTIVATE = 0x0086,
        /// <summary>
        /// 
        /// </summary>
        WM_GETDLGCODE = 0x0087,
        /// <summary>
        /// 
        /// </summary>
        WM_NCMOUSEMOVE = 0x00A0,
        /// <summary>
        /// 
        /// </summary>
        WM_NCLBUTTONDOWN = 0x00A1,
        /// <summary>
        /// 
        /// </summary>
        WM_NCLBUTTONUP = 0x00A2,
        /// <summary>
        /// 
        /// </summary>
        WM_NCLBUTTONDBLCLK = 0x00A3,
        /// <summary>
        /// 
        /// </summary>
        WM_NCRBUTTONDOWN = 0x00A4,
        /// <summary>
        /// 
        /// </summary>
        WM_NCRBUTTONUP = 0x00A5,
        /// <summary>
        /// 
        /// </summary>
        WM_NCRBUTTONDBLCLK = 0x00A6,
        /// <summary>
        /// 
        /// </summary>
        WM_NCMBUTTONDOWN = 0x00A7,
        /// <summary>
        /// 
        /// </summary>
        WM_NCMBUTTONUP = 0x00A8,
        /// <summary>
        /// 
        /// </summary>
        WM_NCMBUTTONDBLCLK = 0x00A9,
        /// <summary>
        /// 
        /// </summary>
        WM_KEYFIRST = 0x0100,
        /// <summary>
        /// 
        /// </summary>
        WM_KEYDOWN = 0x0100,
        /// <summary>
        /// 
        /// </summary>
        WM_KEYUP = 0x0101,
        /// <summary>
        /// 
        /// </summary>
        WM_CHAR = 0x0102,
        /// <summary>
        /// 
        /// </summary>
        WM_DEADCHAR = 0x0103,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSKEYDOWN = 0x0104,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSKEYUP = 0x0105,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSCHAR = 0x0106,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSDEADCHAR = 0x0107,
        /// <summary>
        /// 
        /// </summary>
        WM_KEYLAST = 0x0108,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_STARTCOMPOSITION = 0x010D,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_ENDCOMPOSITION = 0x010E,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_COMPOSITION = 0x010F,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_KEYLAST = 0x010F,
        /// <summary>
        /// 
        /// </summary>
        WM_INITDIALOG = 0x0110,
        /// <summary>
        /// 
        /// </summary>
        WM_COMMAND = 0x0111,
        /// <summary>
        /// 
        /// </summary>
        WM_SYSCOMMAND = 0x0112,
        /// <summary>
        /// 
        /// </summary>
        WM_TIMER = 0x0113,
        /// <summary>
        /// 
        /// </summary>
        WM_HSCROLL = 0x0114,
        /// <summary>
        /// 
        /// </summary>
        WM_VSCROLL = 0x0115,
        /// <summary>
        /// 
        /// </summary>
        WM_INITMENU = 0x0116,
        /// <summary>
        /// 
        /// </summary>
        WM_INITMENUPOPUP = 0x0117,
        /// <summary>
        /// 
        /// </summary>
        WM_MENUSELECT = 0x011F,
        /// <summary>
        /// 
        /// </summary>
        WM_MENUCHAR = 0x0120,
        /// <summary>
        /// 
        /// </summary>
        WM_ENTERIDLE = 0x0121,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORMSGBOX = 0x0132,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLOREDIT = 0x0133,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORLISTBOX = 0x0134,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORBTN = 0x0135,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORDLG = 0x0136,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORSCROLLBAR = 0x0137,
        /// <summary>
        /// 
        /// </summary>
        WM_CTLCOLORSTATIC = 0x0138,
        /// <summary>
        /// 
        /// </summary>
        WM_MOUSEFIRST = 0x0200,
        /// <summary>
        /// 
        /// </summary>
        WM_MOUSEMOVE = 0x0200,
        /// <summary>
        /// 
        /// </summary>
        WM_LBUTTONDOWN = 0x0201,
        /// <summary>
        /// 
        /// </summary>
        WM_LBUTTONUP = 0x0202,
        /// <summary>
        /// 
        /// </summary>
        WM_LBUTTONDBLCLK = 0x0203,
        /// <summary>
        /// 
        /// </summary>
        WM_RBUTTONDOWN = 0x0204,
        /// <summary>
        /// 
        /// </summary>
        WM_RBUTTONUP = 0x0205,
        /// <summary>
        /// 
        /// </summary>
        WM_RBUTTONDBLCLK = 0x0206,
        /// <summary>
        /// 
        /// </summary>
        WM_MBUTTONDOWN = 0x0207,
        /// <summary>
        /// 
        /// </summary>
        WM_MBUTTONUP = 0x0208,
        /// <summary>
        /// 
        /// </summary>
        WM_MBUTTONDBLCLK = 0x0209,
        /// <summary>
        /// 
        /// </summary>
        WM_MOUSELAST = 0x0209,
        /// <summary>
        /// 
        /// </summary>
        WM_PARENTNOTIFY = 0x0210,
        /// <summary>
        /// 
        /// </summary>
        WM_ENTERMENULOOP = 0x0211,
        /// <summary>
        /// 
        /// </summary>
        WM_EXITMENULOOP = 0x0212,
        /// <summary>
        /// 
        /// </summary>
        WM_NEXTMENU = 0x0213,
        /// <summary>
        /// 
        /// </summary>
        WM_SIZING = 0x0214,
        /// <summary>
        /// 
        /// </summary>
        WM_CAPTURECHANGED = 0x0215,
        /// <summary>
        /// 
        /// </summary>
        WM_MOVING = 0x0216,
        /// <summary>
        /// 
        /// </summary>
        WM_POWERBROADCAST = 0x0218,
        /// <summary>
        /// 
        /// </summary>
        WM_DEVICECHANGE = 0x0219,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_SETCONTEXT = 0x0281,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_NOTIFY = 0x0282,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_CONTROL = 0x0283,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_COMPOSITIONFULL = 0x0284,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_SELECT = 0x0285,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_CHAR = 0x0286,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_KEYDOWN = 0x0290,
        /// <summary>
        /// 
        /// </summary>
        WM_IME_KEYUP = 0x0291,
        /// <summary>
        /// 
        /// </summary>
        WM_MDICREATE = 0x0220,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIDESTROY = 0x0221,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIACTIVATE = 0x0222,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIRESTORE = 0x0223,
        /// <summary>
        /// 
        /// </summary>
        WM_MDINEXT = 0x0224,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIMAXIMIZE = 0x0225,
        /// <summary>
        /// 
        /// </summary>
        WM_MDITILE = 0x0226,
        /// <summary>
        /// 
        /// </summary>
        WM_MDICASCADE = 0x0227,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIICONARRANGE = 0x0228,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIGETACTIVE = 0x0229,
        /// <summary>
        /// 
        /// </summary>
        WM_MDISETMENU = 0x0230,
        /// <summary>
        /// 
        /// </summary>
        WM_ENTERSIZEMOVE = 0x0231,
        /// <summary>
        /// 
        /// </summary>
        WM_EXITSIZEMOVE = 0x0232,
        /// <summary>
        /// 
        /// </summary>
        WM_DROPFILES = 0x0233,
        /// <summary>
        /// 
        /// </summary>
        WM_MDIREFRESHMENU = 0x0234,
        /// <summary>
        /// 
        /// </summary>
        WM_CUT = 0x0300,
        /// <summary>
        /// 
        /// </summary>
        WM_COPY = 0x0301,
        /// <summary>
        /// 
        /// </summary>
        WM_PASTE = 0x0302,
        /// <summary>
        /// 
        /// </summary>
        WM_CLEAR = 0x0303,
        /// <summary>
        /// 
        /// </summary>
        WM_UNDO = 0x0304,
        /// <summary>
        /// 
        /// </summary>
        WM_RENDERFORMAT = 0x0305,
        /// <summary>
        /// 
        /// </summary>
        WM_RENDERALLFORMATS = 0x0306,
        /// <summary>
        /// 
        /// </summary>
        WM_DESTROYCLIPBOARD = 0x0307,
        /// <summary>
        /// 
        /// </summary>
        WM_DRAWCLIPBOARD = 0x0308,
        /// <summary>
        /// 
        /// </summary>
        WM_PAINTCLIPBOARD = 0x0309,
        /// <summary>
        /// 
        /// </summary>
        WM_VSCROLLCLIPBOARD = 0x030A,
        /// <summary>
        /// 
        /// </summary>
        WM_SIZECLIPBOARD = 0x030B,
        /// <summary>
        /// 
        /// </summary>
        WM_ASKCBFORMATNAME = 0x030C,
        /// <summary>
        /// 
        /// </summary>
        WM_CHANGECBCHAIN = 0x030D,
        /// <summary>
        /// 
        /// </summary>
        WM_HSCROLLCLIPBOARD = 0x030E,
        /// <summary>
        /// 
        /// </summary>
        WM_QUERYNEWPALETTE = 0x030F,
        /// <summary>
        /// 
        /// </summary>
        WM_PALETTEISCHANGING = 0x0310,
        /// <summary>
        /// 
        /// </summary>
        WM_PALETTECHANGED = 0x0311,
        /// <summary>
        /// 
        /// </summary>
        WM_HOTKEY = 0x0312,
        /// <summary>
        /// 
        /// </summary>
        WM_PRINT = 0x0317,
        /// <summary>
        /// 
        /// </summary>
        WM_PRINTCLIENT = 0x0318,
        /// <summary>
        /// 
        /// </summary>
        WM_HANDHELDFIRST = 0x0358,
        /// <summary>
        /// 
        /// </summary>
        WM_HANDHELDLAST = 0x035F,
        /// <summary>
        /// 
        /// </summary>
        WM_AFXFIRST = 0x0360,
        /// <summary>
        /// 
        /// </summary>
        WM_AFXLAST = 0x037F,
        /// <summary>
        /// 
        /// </summary>
        WM_PENWINFIRST = 0x0380,
        /// <summary>
        /// 
        /// </summary>
        WM_PENWINLAST = 0x038F,
        /// <summary>
        /// 
        /// </summary>
        WM_APP = 0x8000,
        /// <summary>
        /// 
        /// </summary>
        WM_USER = 0x0400
    }
}
