using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.DeveloperToolbox.Translation;

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
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["EnvironmentCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "EnvironmentCaption");
                }
            }
        }

        public string EnvironmentDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["EnvironmentDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "EnvironmentDescription");
                }
            }
        }


        public string FinishCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["FinishCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "FinishCaption");
                }
            }
        }

        public string FinishDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["FinishDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "FinishDescription");
                }
            }
        }


        public string GuiCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["GuiCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "GuiCaption");
                }
            }
        }

        public string GuiDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["GuiDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "GuiDescription");
                }
            }
        }

        public string HostCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["HostCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "HostCaption");
                }
            }
        }

        public string HostDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["HostDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "HostDescription");
                }
            }
        }


        public string LoadCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["LoadCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "LoadCaption");
                }
            }
        }

        public string LoadDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["LoadDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "LoadDescription");
                }
            }
        }


        public string NameCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["NameCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "NameCaption");
                }
            }
        }

        public string NameDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["NameDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "NameDescription");
                }
            }
        }


        public string ProjectCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["ProjectCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "ProjectCaption");
                }
            }
        }

        public string ProjectDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["ProjectDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "ProjectDescription");
                }
            }
        }



        public string SummaryCaption
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["SummaryCaption"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "SummaryCaption");
                }
            }
        }

        public string SummaryDescription
        {
            get
            {
                Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
                if (null != language)
                {
                    var component = language.Components["Project Wizard - Captions"];
                    return component.ControlRessources["SummaryDescription"].Value2;
                }
                else
                {
                    return Translator.GetRessourceValue("ToolboxControls.ProjectWizard.CaptionStrings.txt", Forms.MainForm.Singleton.CurrentLanguageID, "SummaryDescription");
                }
            }
        }

    }
}
