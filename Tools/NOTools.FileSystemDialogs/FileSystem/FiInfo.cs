using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class FiInfo : FileSystemInfo
    {
        internal FiInfo(FileSystemManager parent, string name, string path, long lenght)
            : base(parent, name, path)
        {
            Size = lenght;
            TryGetAssociatedIcon();
        }

        public override bool Exists
        {
            get
            {
                return File.Exists(Path);
            }
        }

        public string ValidatedSize
        {
            get
            {
                return Size.ToString() + " Bytes";
            }
        }

        public long Size { get; private set; }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }

        public override bool IsDrivesLoaded { get; internal set; }

        public override void LoadDrives()
        {
        }

        public override void LoadFiles()
        {
        }

        public override void LoadDirectories()
        {
        }

        private void TryGetAssociatedIcon()
        {
            //Win32.SHFILEINFO shinfoSmall = new Win32.SHFILEINFO();
            //Win32.SHFILEINFO shinfoLarge = new Win32.SHFILEINFO();

            //IntPtr hImgSmall = Win32.SHGetFileInfo(Path, 0, ref shinfoSmall, (int)Marshal.SizeOf(shinfoSmall), Convert.ToInt32(Win32.SHGFI_ICON | Win32.SHGFI_SMALLICON));
            //IntPtr hImgLarge = Win32.SHGetFileInfo(Path, 0, ref shinfoLarge, (int)Marshal.SizeOf(shinfoLarge), Convert.ToInt32(Win32.SHGFI_ICON | Win32.SHGFI_LARGEICON));
            //if (hImgSmall.ToInt32() > 0 && hImgLarge.ToInt32() > 0 &&
            //    shinfoSmall.hIcon.ToInt32() > 0 && shinfoLarge.hIcon.ToInt32() > 0)
            //{
            //    DisplayImageSmall = System.Drawing.Icon.FromHandle(shinfoSmall.hIcon);
            //    DisplayImageLarge = System.Drawing.Icon.FromHandle(shinfoLarge.hIcon);
            //    if (null != DisplayImageSmall && null != DisplayImageLarge)
            //        SupportsDisplayImage = true;            
            //}
        }
    }
}
