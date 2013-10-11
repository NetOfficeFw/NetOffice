using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// All MyComputer for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class MyComputerSettings : DefaultableSettings
    {
        public MyComputerSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {
            DerivedPropertyBag = new PropertyBagCollection<bool>(true, RaisePropertyChanged);
        }

        [Category("Drives"), Description("Get or set unready drives are visible.")]
        public bool ShowUnreadyDrives
        {
            get { return DerivedPropertyBag["ShowUnreadyDrives"]; }
            set { DerivedPropertyBag["ShowUnreadyDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set unkown drives are visible.")]
        public bool ShowUnknownDrives
        {
            get { return DerivedPropertyBag["ShowUnknownDrives"]; }
            set { DerivedPropertyBag["ShowUnknownDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set drives without a root directory are visible.")]
        public bool ShowNoRootDirectoryDrives
        {
            get { return DerivedPropertyBag["ShowNoRootDirectoryDrives"]; }
            set { DerivedPropertyBag["ShowNoRootDirectoryDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set removable drives are visible.")]
        public bool ShowRemovableDrives
        {
            get { return DerivedPropertyBag["ShowRemovableDrives"]; }
            set { DerivedPropertyBag["ShowRemovableDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set fixed drives are visible.")]
        public bool ShowFixedDrives
        {
            get { return DerivedPropertyBag["ShowFixedDrives"]; }
            set { DerivedPropertyBag["ShowFixedDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set network drives are visible.")]
        public bool ShowNetworkDrives
        {
            get { return DerivedPropertyBag["ShowNetworkDrives"]; }
            set { DerivedPropertyBag["ShowNetworkDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set cd drives are visible.")]
        public bool ShowCDRomDrives
        {
            get { return DerivedPropertyBag["ShowCDRomDrives"]; }
            set { DerivedPropertyBag["ShowCDRomDrives"] = value; }
        }

        [Category("Drives"), Description("Get or set ram drives are visible.")]
        public bool ShowRamDrives
        {
            get { return DerivedPropertyBag["ShowRamDrives"]; }
            set { DerivedPropertyBag["ShowRamDrives"] = value; }
        }

        /// <summary>
        /// Dynamic property bag to hold property values
        /// </summary>
        private PropertyBagCollection<bool> DerivedPropertyBag { get; set; }

        private bool IsAllowedDrive(DrvInfo drive)
        {
            if (!drive.IsReady && !ShowUnreadyDrives)
                return false;

            switch (drive.Type)
            {
                case System.IO.DriveType.CDRom:
                    return ShowCDRomDrives;
                case System.IO.DriveType.Fixed:
                    return ShowFixedDrives;
                case System.IO.DriveType.Network:
                    return ShowNetworkDrives;
                case System.IO.DriveType.NoRootDirectory:
                    return ShowNoRootDirectoryDrives;
                case System.IO.DriveType.Ram:
                    return ShowRamDrives;
                case System.IO.DriveType.Removable:
                    return ShowRamDrives;
                case System.IO.DriveType.Unknown:
                    return ShowUnknownDrives;
                default:
                    throw new ArgumentOutOfRangeException("DriveType");
            }
        }

        public override bool HasAllowedSubFolders(FileSystemInfo fsInfo)
        {
            if (!GetRuntimeValue("AllowBrowseFolders"))
                return false;

            bool result = false;
            foreach (var item in fsInfo.Drives)
            {
                if (IsAllowedDrive(item))
                {
                    result = true;
                    return result;
                }
            }

            foreach (var item in fsInfo.Directories)
            {
                if (AllowShowFolder(item))
                {
                    result = true;
                    return result;
                }
            }

            return result;
        }

        public override bool AllowShowDrive(DrvInfo drive)
        {
            return IsAllowedDrive(drive);
        }

        public override string ToString()
        {
            return "MyComputer";
        }
    }
}
