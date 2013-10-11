using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    public interface IApplicationHost
    {
        bool IsCurrentlyVisible(IApplicationControl control);

        bool ShowTime { get; }

        bool ShowMachine { get; }

        bool ShowAppDomain { get; }
    }
}
