using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Controls.Hex;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    /// <summary>
    /// Binary Value Editor
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeBinaryDialogMessageTable.txt")]
    public partial class ChangeBinaryControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ChangeBinaryControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Events

        /// <summary>
        /// User want close the dialog
        /// </summary>
        public event EventHandler Close;

        private void RaiseClose()
        {
            if (null != Close)
                Close(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        /// <summary>
        /// User want proceed or abort edit
        /// </summary>
        public DialogResult DialogResult { get; private set; }

        /// <summary>
        /// New value
        /// </summary>
        public Byte[] Bytes
        {
            get
            {
                DynamicByteProvider provider = hexBox.ByteProvider as DynamicByteProvider;
                return provider.Bytes.ToArray();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set name and value to edit
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">value to edit</param>
        internal void SetArguments(string name, byte[] value)
        {
            DynamicByteProvider provider = new DynamicByteProvider(value);
            hexBox.ByteProvider = provider;
            textBoxName.Text = name;
            hexBox.ByteProvider = provider;
        }

        /// <summary>
        /// Set focus to the edit value
        /// </summary>
        internal void SetFocus()
        {
            hexBox.Focus();
        }

        private string ByteArrayToString(byte[] byteArray)
        {
            string result = "";
            foreach (byte value in byteArray)
            {
                char output = Convert.ToChar(value);
                result += output;
            }
            return result;
        }

        private static byte[] StringToByteArray(string str)
        {
            if (null == str)
                return null;
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetBytes(str);
        }

        #endregion

        #region Trigger

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.OK;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.Cancel;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void ChangeBinaryControl_Resize(object sender, EventArgs e)
        {
            try
            {
                hexBox.BytesPerLine = this.Width / 56;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
