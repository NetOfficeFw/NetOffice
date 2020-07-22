using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Tests.Helpers;

namespace NetOffice.Tests.NetOffice.Events
{
    internal class EventClassStub : TestableComObjectStub
    {
        // backing fields for events in NetOffice must start with _ and end with 'Event'
        private event EventHandler _Custom1Event;
        private event EventHandler _Custom2Event;
        private event EventHandler _Custom3Event;

        public event EventHandler Custom1Event
        {
            add => _Custom1Event += value;
            remove => _Custom1Event -= value;
        }

        public event EventHandler Custom2Event
        {
            add => _Custom2Event += value;
            remove => _Custom2Event -= value;
        }

        public event EventHandler Custom3Event
        {
            add => _Custom3Event += value;
            remove => _Custom3Event -= value;
        }
    }
}
