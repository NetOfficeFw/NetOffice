using System;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Collections.Generic;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Contains information about coresponding localization resource file
    /// </summary>
    public class RessourceTableAttribute : System.Attribute
    {
        /// <summary>
        /// Localization Resource File Adress
        /// </summary>
        public readonly string Address;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="address">localization resource file adress</param>
        public RessourceTableAttribute(string address)
        {
            Address = address;
        }
        
        /// <summary>
        /// Return all names and values for localization based on language
        /// </summary>
        /// <param name="control">instance must have a RessourceTable attribute</param>
        /// <param name="languageID">target language id</param>
        /// <returns>target names and values</returns>
        public static Dictionary<string, string> GetRessourceValues(Control control, int languageID)
        {
            Type type = control.GetType();
            Assembly assembly = type.Assembly;
            object[] obj = type.GetCustomAttributes(typeof(RessourceTableAttribute), false);
            RessourceTableAttribute attrib = obj[0] as RessourceTableAttribute;
            return null;
            //return Translation.Translator.GetTranslateRessources(control, attrib.Address, languageID);
        }

        /// <summary>
        /// Returns all names for localization
        /// </summary>
        /// <param name="type">type of instance(must have RessourceTableAttribute)</param>
        /// <returns>target names</returns>
        public static string[] GetRessourceNames(Type type)
        {
            object[] obj = type.GetCustomAttributes(typeof(RessourceTableAttribute), false);
            RessourceTableAttribute attrib = obj[0] as RessourceTableAttribute;
            Stream stream = type.Assembly.GetManifestResourceStream(type.Assembly.GetName().Name + "." + attrib.Address);
            StreamReader reader = new StreamReader(stream);
            string content = reader.ReadToEnd();
            reader.Dispose();
            stream.Dispose();
            return null;
//            return Translation.Translator.ReadRessourceNames(content);
        }
    }
}
