using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Encapsulate generic addin services 
    /// </summary>
    public abstract class COMAddinBase
    {
        /// <summary>
        /// Creates an instance of th class
        /// </summary>
        public COMAddinBase()
        {
            _creationTime = DateTime.Now;
        }

        /// <summary>
        /// Set in ctor first to measure the time from creation to OnStartupComplete
        /// </summary>
        protected DateTime _creationTime;

        /// <summary>
        /// Host Application Instance
        /// </summary>
        public abstract COMObject AppInstance { get; }

        /// <summary>
        /// Elapsed time in milliseconds from instance creation until OnStartupComplete event
        /// </summary>
        public TimeSpan LoadingTimeElapsed { get; protected set; }
    }
}
