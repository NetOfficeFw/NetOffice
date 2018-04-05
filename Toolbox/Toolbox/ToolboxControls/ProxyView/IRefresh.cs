using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    interface IRefresh : IDisposable
    {
        void RefreshAsync(Action<IRefresh> complete, Control syncRoot);

        bool IsCurrentlyRefresh { get; }
    }
}