using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    public class Translator
    {
        public static string[] ReadRessourceNames(string ressourceContent)
        {
            List<String> list = new List<string>();
            string[] splitArray = ressourceContent.Split(new string[] { "[End]" }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> translateTable = GetTranslateRessources(splitArray, 1031);
            foreach (var item in translateTable)
                list.Add(item.Key);
            return list.ToArray();
        }

        public static void TranslateControl(Control rootControl, string name, string text)
        {
            if (name.Equals("this", StringComparison.InvariantCulture))
                rootControl.Text = text;

            foreach (Control item in rootControl.Controls)
            {
                ToolStrip strip = item as ToolStrip;
                if (null != strip)
                {
                    foreach (ToolStripItem stripItem in strip.Items)
                    {
                        if (stripItem.Name.Equals(name, StringComparison.InvariantCulture))
                        {
                            stripItem.Text = text;
                            return;
                        }
                    }
                }

                if (item.Name.Equals(name, StringComparison.InvariantCulture))
                {
                     item.Text = text;
                     return;
                }
                Control subCtrol = TryGetControl(item, name);
                if (null != subCtrol)
                {
                    subCtrol.Text = text;
                    return;
                }
            }
        }

        public static void AutoTranslateControls(Control control, string componentName, string ressourceFileName, int lanuageID)
        {
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == lanuageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components[componentName];
                Translation.Translator.TranslateControls(control, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(control, ressourceFileName, lanuageID);
            }
        }

        public static void TranslateControls(Control control, ItemCollection strings)
        {
            string caption = "";
            strings.TryGetValue("this", out caption);
            if (!string.IsNullOrEmpty(caption))
                control.Text = caption;

            ILocalizationDesign toolBoxControl = control as ILocalizationDesign;
            if ((null != toolBoxControl) && (null != toolBoxControl.Components))
            {
                foreach (System.ComponentModel.IComponent controlComponent in toolBoxControl.Components.Components)
                {
                    ContextMenuStrip menuStrip = controlComponent as ContextMenuStrip;
                    if (null != menuStrip)
                    {
                        string message = "";
                        strings.TryGetValue(menuStrip.Name, out message);
                        if (!string.IsNullOrEmpty(message))
                            menuStrip.Text = message;

                        foreach (ToolStripItem unkownItem in menuStrip.Items)
                        {
                            ToolStripMenuItem menuItem = unkownItem as ToolStripMenuItem;
                            if (null != menuItem)
                            {
                                message = "";
                                strings.TryGetValue(menuItem.Name, out message);
                                if (!string.IsNullOrEmpty(message))
                                    menuItem.Text = message;
                                ForEachItems(menuItem, strings);
                            }
                        }
                    }
                }
            }

            foreach (Control item in control.Controls)
            {
                ToolStrip toolStrip = item as ToolStrip;
                if (null != toolStrip)
                {
                    string message = "";
                    strings.TryGetValue(toolStrip.Name, out message);
                    if (!string.IsNullOrEmpty(message))
                        toolStrip.Text = message;

                    foreach (ToolStripItem unkownItem in toolStrip.Items)
                    {
                        ToolStripItem menuItem = unkownItem as ToolStripItem;
                        if (null != menuItem)
                        {
                            message = "";
                            strings.TryGetValue(menuItem.Name, out message);
                            if (!string.IsNullOrEmpty(message))
                                menuItem.Text = message;
                            ForEachItems(menuItem, strings);
                        }
                    }
                }
            }

            foreach (Control item in control.Controls)
            {
                string message = "";
                strings.TryGetValue(item.Name, out message);
                if (!string.IsNullOrEmpty(message))
                {
                    RichTextBox box =item as RichTextBox;
                    if (null != box)
                        box.Rtf = message;
                    else
                        item.Text = message;
                }
                ForEachSubControls(item, strings);
            }
        }

        
        public static void TranslateControls(Control control, string ressourceFile, int languageId)
        {
            string ressourceContent = ReadString(ressourceFile);
            string[] splitArray = ressourceContent.Split(new string[] { "[End]" }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> translateTable = GetTranslateRessources(splitArray, languageId);

            ILocalizationDesign toolBoxControl = control as ILocalizationDesign;
            if ((null != toolBoxControl) && (null != toolBoxControl.Components))
            {
                foreach (System.ComponentModel.IComponent controlComponent in toolBoxControl.Components.Components)
                {
                    ContextMenuStrip menuStrip = controlComponent as ContextMenuStrip;
                    if (null != menuStrip)
                    {
                        string message = "";
                        translateTable.TryGetValue(menuStrip.Name, out message);
                        if (!string.IsNullOrEmpty(message))
                            menuStrip.Text = message;

                        foreach (ToolStripItem unkownItem in menuStrip.Items)
                        {
                            ToolStripMenuItem menuItem = unkownItem as ToolStripMenuItem;
                            if (null != menuItem)
                            { 
                                message = "";
                                translateTable.TryGetValue(menuItem.Name, out message);
                                if (!string.IsNullOrEmpty(message))
                                    menuItem.Text = message;
                                ForEachItems(menuItem, translateTable);
                            }
                        }
                    }
                }
            }

            string caption = "";
            translateTable.TryGetValue("this", out caption);
            if (!string.IsNullOrEmpty(caption))
                control.Text = caption;

            foreach (Control item in control.Controls)
            {
                string message = "";
                translateTable.TryGetValue(item.Name, out message);
                if (!string.IsNullOrEmpty(message))
                    item.Text = message;
                ForEachSubControls(item, translateTable);
            }
        }

        public static string GetRessourceValue(string ressourceFile, int languageId, string ressourceName)
        {
            string ressourceContent = ReadString(ressourceFile);
            string[] splitArray = ressourceContent.Split(new string[] { "[End]" }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> translateTable = GetTranslateRessources(splitArray, languageId);
            var res = translateTable.Where(n => n.Key == ressourceName).FirstOrDefault();
            if (null != res.Key)
                return res.Value;
            else
                return null;
        }
    
        public static Dictionary<string, string> GetTranslateRessources(Control control, string ressourceFile, int languageId)
        {
            string ressourceContent = ReadString(ressourceFile);
            string[] splitArray = ressourceContent.Split(new string[] { "[End]" }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> translateTable = GetTranslateRessources(splitArray, languageId, control as ILocalizationReplaceProvider);
            return translateTable;
        }

        internal static string TryGetControlText(Control rootControl, string name)
        {
            if (name.Equals("this", StringComparison.InvariantCulture))
                return rootControl.Text;
            foreach (Control item in rootControl.Controls)
            {
                ToolStrip strip = item as ToolStrip;
                if (null != strip)
                {
                    foreach (ToolStripItem stripItem in strip.Items)
                    {
                        if (stripItem.Name.Equals(name, StringComparison.InvariantCulture))
                            return stripItem.Text;
                    }
                }

                if (item.Name.Equals(name, StringComparison.InvariantCulture))
                    return item.Text;
                Control subCtrol = TryGetControl(item, name);
                if (null != subCtrol)
                    return subCtrol.Text;
            }
            return null;
        }

        internal static Control TryGetControl(Control rootControl, string name)
        {
            if(name.Equals("this", StringComparison.InvariantCulture))
                return rootControl;
            foreach (Control item in rootControl.Controls)
            {
                ToolStrip strip = item as ToolStrip;
                if (null != strip)
                {
                    foreach (ToolStripItem stripItem in strip.Items)
                    {
                        if (stripItem.Name.Equals(name, StringComparison.InvariantCulture))
                            return item;
                    }
                }

                if (item.Name.Equals(name, StringComparison.InvariantCulture))
                    return item;
                Control subCtrol = TryGetControl(item, name);
                if (null != subCtrol)
                    return subCtrol;
            }
            return null;
        }

        private static void ForEachItems(ToolStripItem item, ItemCollection translateTable)
        {
            // dumy
        }

        private static void ForEachItems(ToolStripMenuItem item, ItemCollection translateTable)
        {
            foreach (ToolStripMenuItem subItem in item.DropDownItems)
            {
                string message = "";
                translateTable.TryGetValue(subItem.Name, out message);
                if (!string.IsNullOrEmpty(message))
                    subItem.Text = message;
                ForEachItems(subItem, translateTable);
            }
        }

        private static void ForEachItems(ToolStripMenuItem item, Dictionary<string, string> translateTable)
        {
            foreach (ToolStripMenuItem subItem in item.DropDownItems)
            {
                string message = "";
                translateTable.TryGetValue(subItem.Name, out message);
                if (!string.IsNullOrEmpty(message))
                    subItem.Text = message;
                ForEachItems(subItem, translateTable);
            }
        }

        private static void ForEachSubControls(Control item, Dictionary<string, string> translateTable)
        {
            foreach (Control subItem in item.Controls)
            {
                string message = "";
                translateTable.TryGetValue(subItem.Name, out message);
                if (!string.IsNullOrEmpty(message))
                    subItem.Text = message;
                ForEachSubControls(subItem, translateTable);
            }
        }

        private static void ForEachSubControls(Control item, ItemCollection translateTable)
        {
            foreach (Control subItem in item.Controls)
            {
                string message = "";
                translateTable.TryGetValue(subItem.Name, out message);
                if (!string.IsNullOrEmpty(message))
                {
                    RichTextBox box = subItem as RichTextBox;
                    if (null != box && !(box is Controls.Text.AdvRichTextBox))
                        box.Rtf = message;
                    else
                        subItem.Text = message;
                }
                ForEachSubControls(subItem, translateTable);
            }
        }

        private static Dictionary<string, string> GetTranslateRessources(string[] splitArray, int languageId, ILocalizationReplaceProvider provider = null)
        {
            Dictionary<string, string> resultDictionary = new Dictionary<string, string>();

            foreach (string item in splitArray)
            {
                string[] lines = item.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    if ("[" + languageId.ToString() + "]" == line.Trim())
                    {
                        AddToDictionary(resultDictionary, lines, provider);
                        return resultDictionary;
                    }
                }
            }

            return resultDictionary;
            //throw new IndexOutOfRangeException(languageId.ToString() + " not found.");
        }

        private static void AddToDictionary(Dictionary<string, string> resultDictionary, string[] lines, ILocalizationReplaceProvider provider = null)
        {
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i];
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    int position = line.IndexOf("=", StringComparison.InvariantCultureIgnoreCase);
                    string name = line.Substring(0, position - 1).Trim();
                    string value = line.Substring(position + 1).Trim();

                    if (null != provider)
                    {
                        int startIndex = value.IndexOf("{0:$", 0);
                        if (startIndex > -1)
                        {
                            int endIndex = value.IndexOf("}", startIndex + 1);
                            if (endIndex > -1)
                            {
                                string marker = value.Substring(startIndex, endIndex - startIndex +1);
                                string replaceContent = provider.Replace(marker);
                                value = value.Replace(marker, replaceContent);
                            }
                        }
                    }

                    resultDictionary.Add(name, value);
                }
            }
        }

        public static string ReadString(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            System.IO.StreamReader textStreamReader = null;
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                ressourcePath = assemblyName + "." + ressourcePath;
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
            finally
            {
                if (null != textStreamReader)
                    textStreamReader.Close();
                if (null != ressourceStream)
                    ressourceStream.Close();
            }
        }
    }
}
