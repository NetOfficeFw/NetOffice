using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

namespace NetOffice.Tools
{
    /// <summary>
    /// Encapsulate generic addin services 
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddinBase
    {
        /// <summary>
        /// Set in ctor first to measure the time from creation to OnStartupComplete
        /// </summary>
        protected DateTime _creationTime;

        /// <summary>
        /// Type cache field
        /// </summary>
        private Type _type;

        /// <summary>
        /// Static visual styles lock
        /// </summary>
        private static object _lock = new object();

        /// <summary>
        /// Creates an instance of th class
        /// </summary>
        public COMAddinBase()
        {
            _creationTime = DateTime.Now;
            EnableVisualStyles();
        }

        /// <summary>
        /// Host Application Instance
        /// </summary>
        public abstract ICOMObject AppInstance { get; }

        /// <summary>
        /// Current asscociated Core
        /// </summary>
        public abstract Core Factory { get; }

        /// <summary>
        /// Elapsed time in milliseconds from instance creation until OnStartupComplete event
        /// </summary>
        public TimeSpan LoadingTimeElapsed { get; protected set; }

        /// <summary>
        /// Type Information of the instance
        /// </summary>
        public Type Type
        {
            get
            {
                if (null == _type)
                    _type = GetType();
                return _type;
            }
        }

        /// <summary>
        /// Instance managed root com objects
        /// </summary>
        [System.ComponentModel.Browsable(false), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        public abstract IEnumerable Roots { get; protected set; }

        /// <summary>
        /// Call System.Windows.Forms.Application.EnableVisualStyles
        /// </summary>
        protected internal virtual void EnableVisualStyles()
        {
            lock (_lock)
            {
                if (System.Windows.Forms.Application.VisualStyleState == System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled)
                    System.Windows.Forms.Application.EnableVisualStyles();
            }
        } 
    }
}
