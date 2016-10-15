using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;

namespace ProxyView
{
    internal class Entry
    {
        public Entry()
        {

        }

        public Entry(object underlying, string id, string caption, string name, string component, string libraryID,
            string proceesID, ProxyInformation.ProcessElevation elevated)
        {
            Underlying = underlying;
            ID = id == Guid.Empty.ToString() ? "<Unknown>" : id;
            Caption = String.IsNullOrWhiteSpace(caption) ? "<Unknown>" : caption;
            Name = String.IsNullOrWhiteSpace(name) ? "<Unknown>" : name;
            Component = String.IsNullOrWhiteSpace(component) ? "<Unknown>" : component;
            Library = libraryID == String.Empty ? "<Unknown>" : libraryID;
            ProcessID = proceesID;
            Elevated = elevated;
        }

        public string Caption { get; private set; }

        public string Name { get; private set; }

        public string Component { get; private set; }

        public string ID { get; private set; }

        public string Library { get; private set; }

        public string ProcessID { get; private set; }

        public ProxyInformation.ProcessElevation Elevated { get; private set; }
    
        [Browsable(false)]
        public object Underlying { get; private set; }
    }
}
