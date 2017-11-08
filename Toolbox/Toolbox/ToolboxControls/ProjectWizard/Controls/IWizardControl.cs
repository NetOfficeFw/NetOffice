using System;
using System.Xml;
using System.Windows.Forms; 
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    public delegate void ReadyStateChangedHandler(Control sender);

    public enum ImageType
    {
        Question = 0,
        Finish = 1
    }

    public interface IWizardControl
    {       
        event ReadyStateChangedHandler ReadyStateChanged;

        bool IsReadyForNextStep { get; }
         
        string Caption { get; }

        string Description { get; }

        ImageType Image { get; }

        void Activate();

        void Deactivate();

        XmlDocument SettingsDocument { get; }

        string[] GetSettingsSummary();

        void KeyDown(KeyEventArgs e);
    }
}