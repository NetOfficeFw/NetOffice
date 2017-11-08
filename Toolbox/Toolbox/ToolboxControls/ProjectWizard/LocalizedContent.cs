using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard
{
    internal class LocalizedContent
    {
        public string StepProgress
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "StepProgress");
            }
        }

        public string Completed
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Completed");
            }
        }

        public string Yes
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Yes");
            }
        }

        public string No
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "No");
            }
        }

        public string AddinStartup
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "AddinOption1");
            }
        }

        public string AddinOnDemand
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "AddinOption2");
            }
        }

        public string AddinNotAutomaticaly
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "AddinOption3");
            }
        }

        public string AddinFirstTime
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "AddinOption4");
            }
        }

        public string Registry
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Registry");
            }
        }

        public string RegistryCurrentUser
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "CurrentUser");
            }
        }

        public string RegistryLocalMachine
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "LocalMachine");
            }
        }

        public string LoadBehavior
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "LoadBehavior");
            }
        }

        public string Language
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Language");
            }
        }

        public string Runtime
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Runtime");
            }
        }

        public string Applications
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "Applications");
            }
        }

        public string ProjectType
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "ProjectType");
            }
        }

        public string ProjectFolder
        {
            get
            {
                return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", "ProjectFolder");
            }
        }
    }
}
