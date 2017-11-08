using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    interface IRefresh : IDisposable
    {
        void Refresh();

        bool IsCurrentlyRefresh { get; }
    }
}
