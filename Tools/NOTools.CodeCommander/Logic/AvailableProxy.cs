using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CodeCommander.Logic
{
    /// <summary>
    /// means a current available proxy from the host application
    /// </summary>
    internal class AvailableProxy
    {
        /// <summary>
        /// creates an instance of the class
        /// </summary>
        /// <param name="id">id of the proxy</param>
        /// <param name="name">name of the proxy</param>
        internal AvailableProxy(int id, string name)
        {
            ID = id;
            Name = name;
        }

        /// <summary>
        /// Id of the proxy
        /// </summary>
        public int ID { get; private set; }

        /// <summary>
        /// Name of the proxy
        /// </summary>
        public string Name { get; private set; }
    }
}
