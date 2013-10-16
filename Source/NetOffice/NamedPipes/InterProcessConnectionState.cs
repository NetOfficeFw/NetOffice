using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.NamedPipes
{
    internal enum InterProcessConnectionState
    {
        NotSet = 0,
        Error = 1,
        Creating = 2,
        Created = 3,
        WaitingForClient = 4,
        ConnectedToClient = 5,
        ConnectingToServer = 6,
        ConnectedToServer = 7,
        Reading = 8,
        ReadData = 9,
        Writing = 10,
        WroteData = 11,
        Flushing = 12,
        FlushedData = 13,
        Disconnecting = 14,
        Disconnected = 15,
        Closing = 16,
        Closed = 17,
    }
}
