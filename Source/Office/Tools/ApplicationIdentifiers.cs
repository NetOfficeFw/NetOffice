using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Contains the Application Interface ID's.
    /// This is useful to determine a COM proxy represents an application.
    /// </summary>
    public static class ApplicationIdentifiers
    {
        /// <summary>
        /// Represents Compare Result
        /// </summary>
        public enum ApplicationType
        {
            /// <summary>
            /// Not a known application
            /// </summary>
            None = 0,

            /// <summary>
            /// Excel
            /// </summary>
            Excel = 1,

            /// <summary>
            /// Word
            /// </summary>
            Word = 2,

            /// <summary>
            /// Outlook
            /// </summary>
            Outlook = 3,

            /// <summary>
            /// PowerPoint
            /// </summary>
            PowerPoint = 4,

            /// <summary>
            /// Access
            /// </summary>
            Access = 5,

            /// <summary>
            /// Visio
            /// </summary>
            Visio = 6,

            /// <summary>
            /// MS Project
            /// </summary>
            MS_Project = 7,

            /// <summary>
            /// Publisher
            /// </summary>
            Publisher = 8
        }

        static ApplicationIdentifiers()
        {
            Excel = new Guid("000208D5-0000-0000-C000-000000000046");
            Word = new Guid("00020970-0000-0000-C000-000000000046");
            Outlook = new Guid("00063001-0000-0000-C000-000000000046");
            PowerPoint = new Guid("91493442-5A91-11CF-8700-00AA0060263B");
            Access = new Guid("68CCE6C0-6129-101B-AF4E-00AA003F0F07");
            Visio = new Guid("000D0700-0000-0000-C000-000000000046");
            MS_Project = new Guid("00020AFF-0000-0000-C000-000000000046");
            Publisher = new Guid("0002123E-0000-0000-C000-000000000046");
        }

        /// <summary>
        /// 000208D5-0000-0000-C000-000000000046
        /// </summary>
        public static Guid Excel {get; private set;}

        /// <summary>
        /// 00020970-0000-0000-C000-000000000046
        /// </summary>
        public static Guid Word { get; private set; }

        /// <summary>
        /// 00063001-0000-0000-C000-000000000046
        /// </summary>
        public static Guid Outlook { get; private set; }

        /// <summary>
        /// 91493442-5A91-11CF-8700-00AA0060263B
        /// </summary>
        public static Guid PowerPoint { get; private set; }
         
        /// <summary>
        /// 68CCE6C0-6129-101B-AF4E-00AA003F0F07
        /// </summary>
        public static Guid Access { get; private set; }
         
        /// <summary>
        /// 000D0700-0000-0000-C000-000000000046
        /// </summary>
        public static Guid Visio { get; private set; }

        /// <summary>
        /// 00020AFF-0000-0000-C000-000000000046
        /// </summary>
        public static Guid MS_Project { get; private set; }

        /// <summary>
        /// 0002123E-0000-0000-C000-000000000046
        /// </summary>
        public static Guid Publisher { get; private set; }

        /// <summary>
        /// Compare the id with application interface id's.
        /// (Typical you got with comProxy.GetType().GUID)
        /// </summary>
        /// <param name="id">given id as any</param>
        /// <returns>application kind or none</returns>
        public static ApplicationType IsApplication(Guid id)
        {
            if (id == Excel)
                return ApplicationType.Excel;
            if (id == Word)
                return ApplicationType.Word;
            if (id == Outlook)
                return ApplicationType.Outlook;
            if (id == PowerPoint)
                return ApplicationType.PowerPoint;
            if (id == Access)
                return ApplicationType.Access;
            if (id == Visio)
                return ApplicationType.Visio;
            if (id == MS_Project)
                return ApplicationType.MS_Project;
            if (id == Publisher)
                return ApplicationType.Publisher;
            return ApplicationType.None;
        }

        /// <summary>
        /// Converts an ApplicationType value to string that can also be used as registry key name
        /// </summary>
        /// <param name="applicationType">target value to convert</param>
        /// <returns>System.String</returns>
        public static string ConvertApplicationType(ApplicationType applicationType)
        {
            return Enum.GetName(typeof(ApplicationType), applicationType).Replace("_", "");
        }
    }
}
