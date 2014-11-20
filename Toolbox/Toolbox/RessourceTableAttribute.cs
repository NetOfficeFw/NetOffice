using System;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    public class RessourceTableAttribute : System.Attribute
    {
        public readonly string Address;

        public RessourceTableAttribute(string address)
        {
            Address = address;
        }
        
        public static Dictionary<string, string> GetRessourceValues(Control control, int languageID)
        {
            Type type = control.GetType();
            Assembly assembly = type.Assembly;
            object[] obj = type.GetCustomAttributes(typeof(RessourceTableAttribute), false);
            RessourceTableAttribute attrib = obj[0] as RessourceTableAttribute;

            return Translation.Translator.GetTranslateRessources(control, attrib.Address, languageID);
        }

        public static string[] GetRessourceNames(Type type)
        {
            object[] obj = type.GetCustomAttributes(typeof(RessourceTableAttribute), false);
            RessourceTableAttribute attrib = obj[0] as RessourceTableAttribute;
            Stream stream = type.Assembly.GetManifestResourceStream(type.Assembly.GetName().Name + "." + attrib.Address);
            StreamReader reader = new StreamReader(stream);
            string content = reader.ReadToEnd();
            reader.Dispose();
            stream.Dispose();
            return Translation.Translator.ReadRessourceNames(content);
        }
    }
}
