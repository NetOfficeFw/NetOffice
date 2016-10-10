using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProxyView
{
    interface IRefresh : IDisposable
    {
        void Refresh();

        bool IsCurrentlyRefresh { get; }
    }
}
