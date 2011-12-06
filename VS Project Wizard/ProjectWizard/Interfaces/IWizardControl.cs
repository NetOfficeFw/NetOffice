using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace NetOffice.ProjectWizard
{
    internal interface IWizardControl
    {
        event ReadyStateChangedHandler ReadyStateChanged;
        bool IsReadyForNextStep { get; }
         
        string Caption { get; }
        string Description { get; }
        ImageType Image { get; }

        void Activate();
        XmlDocument SettingsDocument { get; }
        string[] GetSettingsSummary();
    }

    internal delegate void ReadyStateChangedHandler(IWizardControl sender);
}
