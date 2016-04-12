using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    class WatchNotify
    {
        #region Fields

        NotifyIcon _trayIcon;
        WatchController _parent;

        #endregion

        #region Events

        public event EventHandler MessageFired;
    
        #endregion

        #region Construction

        internal WatchNotify(WatchController parent)
        {
            _parent = parent;
            _trayIcon = new NotifyIcon(new System.ComponentModel.Container());
        }
        
        #endregion
        
        #region Methods

        public void ShowNotification(AddinKey item, NotifyKind notfiyKind, RegistryChangeInfo changeInfo)
        {
            if (_parent.FirstRun)
                return;
            string message = GetMessage(notfiyKind);
            switch (notfiyKind)
            {
                case NotifyKind.AddinLoadBehaviorRestored:
                    AddinValueValueRestoredInfo restoredInfo = (AddinValueValueRestoredInfo)changeInfo;
                    message = string.Format(message, restoredInfo.RootKey, restoredInfo.KeyPath, restoredInfo.ValueName, restoredInfo.RestoredValue, restoredInfo.OldValue, Environment.NewLine);
                    break;
                case NotifyKind.AddinValueNameIsChanged:
                    AddinValueNameChangedInfo nameInfo = (AddinValueNameChangedInfo)changeInfo;
                    message = string.Format(message, nameInfo.RootKey, nameInfo.KeyPath, nameInfo.OldValueName, nameInfo.NewValueName, Environment.NewLine);
                    break;
                case NotifyKind.AddinValueKindIsChanged:
                    AddinValueKindChangedInfo kindInfo = (AddinValueKindChangedInfo)changeInfo;
                    message = string.Format(message, kindInfo.RootKey, kindInfo.KeyPath, kindInfo.ValueName, kindInfo.OldValueKind, kindInfo.NewValueKind, Environment.NewLine);
                    break;
                case NotifyKind.AddinValueIsChanged:
                    AddinValueValueChangedInfo valueInfo = (AddinValueValueChangedInfo)changeInfo;
                    message = string.Format(message, valueInfo.RootKey, valueInfo.KeyPath, valueInfo.ValueName, valueInfo.OldValue, valueInfo.NewValue, Environment.NewLine);
                    break;
                case NotifyKind.AddinValuesIncrement:
                    AddinValuesIncrementInfo incrementInfo = (AddinValuesIncrementInfo)changeInfo;
                    message = string.Format(message, incrementInfo.RootKey, incrementInfo.KeyPath, incrementInfo.ValueName, Environment.NewLine);
                    break;
                case NotifyKind.AddinValuesDecrement:
                    AddinValuesDecrementInfo decrementInfo = (AddinValuesDecrementInfo)changeInfo;
                    message = string.Format(message, decrementInfo.RootKey, decrementInfo.KeyPath, decrementInfo.ValueName, Environment.NewLine);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(notfiyKind.ToString() + " is not valid in this context");
            }

            if (_parent.NotifyType == NotificationType.MessageBox)
                MessageBox.Show(message, "NetOffice.DeveloperToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                _trayIcon.ShowBalloonTip(2000, message, "NetOffice.DeveloperToolbox " + notfiyKind.ToString(), ToolTipIcon.Info);

            if (null != MessageFired)
                MessageFired(message, new EventArgs());
        }

        public void ShowNotification(DisabledKey item, NotifyKind notfiyKind, RegistryChangeInfo changeInfo)
        {
            if (_parent.FirstRun)
                return;
            string message = GetMessage(notfiyKind);
            switch (notfiyKind )
            {
                case NotifyKind.DisabledItemNew:
                    NewDeactivatedElementInfo newInfo = (NewDeactivatedElementInfo)changeInfo;
                    message = string.Format(message, newInfo.OfficeProductVersion, newInfo.Name, Environment.NewLine);
                    break;
                case NotifyKind.DisabledItemDelete:
                    DeleteDeactivatedElementInfo deleteInfo = (DeleteDeactivatedElementInfo)changeInfo;
                    message = string.Format(message, deleteInfo.OfficeProductVersion, deleteInfo.Name, Environment.NewLine);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(notfiyKind.ToString() + " is not valid in this context");
            }

            if (_parent.NotifyType == NotificationType.MessageBox)
                MessageBox.Show(message, "NetOffice.DeveloperToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                _trayIcon.ShowBalloonTip(2000, message, "NetOffice.DeveloperToolbox " + notfiyKind.ToString(), ToolTipIcon.Info);
            
            if (null != MessageFired)
                MessageFired(message, new EventArgs());
        }

        public void ShowNotification(AddinsKey item, NotifyKind notfiyKind, RegistryChangeInfo changeInfo)
        {
            if (_parent.FirstRun)
                return;
      
            string message = GetMessage(notfiyKind);
            switch (notfiyKind)
            {
                case NotifyKind.AddinSubKeysIncrement:
                    AddinSubkeysIncrementInfo newInfo = (AddinSubkeysIncrementInfo)changeInfo;
                    message = String.Format(message, newInfo.RootKey, newInfo.KeyPath, newInfo.KeyName, Environment.NewLine);
                    break;
                case NotifyKind.AddinSubKeysDecrement:
                    AddinSubkeysDecrementInfo deleteInfo = (AddinSubkeysDecrementInfo)changeInfo;
                    message = String.Format(message, deleteInfo.RootKey, deleteInfo.KeyPath, deleteInfo.KeyName, Environment.NewLine);
                    break;
                case NotifyKind.AddinSubKeyNameChanged:
                    AddinSubkeyNameChangedInfo nameInfo = (AddinSubkeyNameChangedInfo)changeInfo;
                    message = String.Format(message, nameInfo.RootKey, nameInfo.KeyPath, nameInfo.OldKeyName, nameInfo.NewKeyName, Environment.NewLine);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(notfiyKind.ToString() + " is not valid in this context");            }

            if (_parent.NotifyType == NotificationType.MessageBox)
                MessageBox.Show(message, "NetOffice.DeveloperToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                _trayIcon.ShowBalloonTip(2000, message, "NetOffice.DeveloperToolbox " + notfiyKind.ToString(), ToolTipIcon.Info);

            if (null != MessageFired)
                MessageFired(message, new EventArgs());
        }
        
        #endregion

        #region Private Methods

        private string GetMessage(NotifyKind notfiyKind)
        {
            string id = notfiyKind.ToString();
            int languageId = _parent.ActiveLanguageID;
            string message = GetMessage(id, languageId);
            if (null == message)
                throw new ArgumentException(notfiyKind.ToString() + " not found.");
            return message;
        }

        private string GetMessage(string key, int languageId)
        {
            string ressourceString = ReadString("AddinGuard.Messages.txt");
            string[] splitArray = ressourceString.Split(new string[] { "[End]" }, StringSplitOptions.RemoveEmptyEntries);

            Dictionary<string, string> transLateTable = GetTranslateRessources(splitArray, languageId);

            string message = "";
            transLateTable.TryGetValue(key, out message);

            return message;
        }

        private Dictionary<string, string> GetTranslateRessources(string[] splitArray, int languageId)
        {
            Dictionary<string, string> resultDictionary = new Dictionary<string, string>();

            foreach (string item in splitArray)
            {
                string[] lines = item.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    if ("[" + languageId.ToString() + "]" == line.Trim())
                    {
                        AddToDictionary(resultDictionary, lines);
                        return resultDictionary;
                    }
                }
            }

            throw new IndexOutOfRangeException(languageId.ToString() + " not found.");
        }

        private void AddToDictionary(Dictionary<string, string> resultDictionary, string[] lines)
        {
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i];
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    int position = line.IndexOf("=", StringComparison.InvariantCultureIgnoreCase);
                    string name = line.Substring(0, position - 1).Trim();
                    string value = line.Substring(position + 1);
                    resultDictionary.Add(name, value);
                }
            }
        }

        private static string ReadString(string ressourcePath)
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

        #endregion
    }
}
