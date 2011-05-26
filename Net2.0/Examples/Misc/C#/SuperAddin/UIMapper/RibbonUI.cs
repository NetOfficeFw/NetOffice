using System;
using System.Collections.Generic;
using Microsoft.Win32;
using Extensibility;
using System.Runtime.InteropServices;
using System.Text;

using System.Runtime.CompilerServices;

namespace SuperAddin.UIMapper
{
    /// <summary>
    /// handles ribbon ui
    /// </summary>
    internal class RibbonUI
    {
        #region Fields

        AddinUI _parent;
        
        #endregion

        #region Construction

        public RibbonUI(AddinUI parent)
        {
            _parent = parent;
        }

        #endregion

        #region Ribbon Methods

        public string GetCustomUI(string RibbonID)
        {
            _parent.RibbonIsActive = true;
            return ReadTextFileFromRessource("RibbonUI.xml");
        }

        public void OnAction(IRibbonControl control)
        {
            try
            {
                _parent.RaiseButtonClick(new ButtonClickArgs(control));
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnAction.", throwedException);
            }
        }
        
        #endregion

        #region Static Helper
 
        private static string ReadTextFileFromRessource(string fileName)
        {
            fileName = "SuperAddin.UIMapper." + fileName;

            System.IO.Stream ressourceStream;
            System.IO.StreamReader textStreamReader;
            try
            {
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                ressourceStream.Close();
                textStreamReader.Close();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
        }
        
        #endregion
    }
}
