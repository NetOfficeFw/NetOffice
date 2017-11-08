using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard
{
    internal class LocalizedCaptions
    {
        private Type _thisType;

        public LocalizedCaptions()
        {
            _thisType = this.GetType();
        }

        public string GetCaption(IWizardControl ctrl)
        {
            if (null == ctrl)
                return "<Error>";
            string name = ctrl.GetType().Name;
            name = name.Substring(0, name.Length - "Control".Length);
            name += "Caption";
            string result = _thisType.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, this, null) as string;
            return result;
        }

        public string GetDescription(IWizardControl ctrl)
        {
            if (null == ctrl)
                return "<Error>";
            string name = ctrl.GetType().Name;
            name = name.Substring(0, name.Length - "Control".Length);
            name += "Description";
            string result = _thisType.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, this, null) as string;
            return result;
        }

        public string EnvironmentCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "EnvironmentCaption");
            }
        }

        public string EnvironmentDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "EnvironmentDescription");
            }
        }

        public string FinishCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "FinishCaption");
            }
        }

        public string FinishDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "FinishDescription");
            }
        }

        public string GuiCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "GuiCaption");
            }
        }

        public string GuiDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "GuiDescription");
            }
        }

        public string HostCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "HostCaption");
            }
        }

        public string HostDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "HostDescription");
            }
        }

        public string LoadCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "LoadCaption");
            }
        }

        public string LoadDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "LoadDescription");
            }
        }

        public string NameCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "NameCaption");
            }
        }

        public string NameDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "NameDescription");
            }
        }


        public string ProjectCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "ProjectCaption");
            }
        }

        public string ProjectDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "ProjectDescription");
            }
        }
        
        public string SummaryCaption
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt",  "SummaryCaption");
            }
        }

        public string SummaryDescription
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", "SummaryDescription");
            }
        }

    }
}