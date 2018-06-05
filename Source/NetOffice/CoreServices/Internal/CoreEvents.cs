using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices.Internal
{
    internal class CoreEvents : ICoreEvents
    {
        private object _thisLock = new object();
        private List<SinkHelper> _pointList = new List<SinkHelper>();

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal CoreEvents(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }
       
        /// <summary>
        /// Count of current opened event bridges
        /// </summary>
        public int Count
        {
            get
            {
                return _pointList.Count;
            }
        }

        /// <summary>
        /// Dispose all active event bridges
        /// </summary>
        public void DisposeAllEventBridges()
        {
            lock (_thisLock)
            {
                foreach (SinkHelper point in _pointList)
                    point.RemoveEventBinding(false);
                _pointList.Clear();
            }
        }

        /// <summary>
        /// Add sink helper to the factory sinkhelper table
        /// </summary>
        /// <param name="point">sink helper as any</param>
        internal void AddEventBridge(SinkHelper point)
        {
            lock (_thisLock)
            {
                _pointList.Add(point);
            }
        }

        /// <summary>
        /// Removes sink helper from factory sinkhelper table.
        /// The method doesnt dispose the argument.
        /// </summary>
        /// <param name="point">sink helper as any</param>
        /// <returns>true if removed, otherwise false</returns>
        internal bool RemoveEventBridge(SinkHelper point)
        {
            lock (_thisLock)
            {
                return _pointList.Remove(point);
            }
        }
    }
}
