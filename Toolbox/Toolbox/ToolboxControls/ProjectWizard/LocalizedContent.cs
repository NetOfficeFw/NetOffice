using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.DeveloperToolbox.Translation;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard
{
    internal class LocalizedContent
    {
        public string StepProgress
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["StepProgress"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "StepProgress");
                }
            }
        }

        public string Completed
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Completed"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Completed");
                }
            }
        }

        public string Yes
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Yes"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Yes");
                }
            }
        }

        public string No
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["No"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "No");
                }
            }
        }

        public string AddinStartup
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["AddinOption1"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "AddinOption1");
                }
            }
        }

        public string AddinOnDemand
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["AddinOption2"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "AddinOption2");
                }
            }
        }

        public string AddinNotAutomaticaly
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["AddinOption3"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "AddinOption3");
                }
            }
        }

        public string AddinFirstTime
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["AddinOption4"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "AddinOption4");
                }
            }
        }

        public string Registry
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Registry"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Registry");
                }
            }
        }

        public string RegistryCurrentUser
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["CurrentUser"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "CurrentUser");
                }
            }
        }

        public string RegistryLocalMachine
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["LocalMachine"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "LocalMachine");
                }
            }
        }

        public string LoadBehavior
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["LoadBehavior"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "LoadBehavior");
                }
            }
        }

        public string Language
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Language"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Language");
                }
            }
        }

        public string Runtime
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Runtime"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Runtime");
                }
            }
        }

        public string Applications
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["Applications"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "Applications");
                }
            }
        }

        public string ProjectType
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["ProjectType"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "ProjectType");
                }
            }
        }

        public string ProjectFolder
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Messages"];
                    return component.ControlRessources["ProjectFolder"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.LocalizationStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "ProjectFolder");
                }
            }
        }
    }
}
