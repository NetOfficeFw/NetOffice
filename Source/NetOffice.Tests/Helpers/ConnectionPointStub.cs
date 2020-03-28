using System;
using System.Runtime.InteropServices.ComTypes;

namespace NetOffice.Tests.Helpers
{
    public class ConnectionPointStub : IConnectionPoint
    {
        public void GetConnectionInterface(out Guid pIID)
        {
            pIID = Guid.Empty;
        }

        public void GetConnectionPointContainer(out IConnectionPointContainer ppCPC)
        {
            ppCPC = null;
        }

        public void Advise(object pUnkSink, out int pdwCookie)
        {
            pdwCookie = 0xCDCDCD;
        }

        public void Unadvise(int dwCookie)
        {
        }

        public void EnumConnections(out IEnumConnections ppEnum)
        {
            ppEnum = null;
        }
    }
}
