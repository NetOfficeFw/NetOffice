using System;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// PropertyGrid Dialog
    /// </summary>
    public partial class PropertyGridDialog : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public PropertyGridDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="selectedObject">object as any</param>
        public PropertyGridDialog(object selectedObject)
        {
            InitializeComponent();
            propertyGrid1.SelectedObject = selectedObject;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Show modal dialog
        /// </summary>
        /// <param name="owner">owner as any</param>
        /// <param name="selectedObject">object as any</param>
        /// <param name="text">optional dialog caption</param>
        public static void ShowForm(IWin32Window owner, object selectedObject, string text = null)
        {
            PropertyGridDialog dialog = new PropertyGridDialog(selectedObject);
            if (!String.IsNullOrWhiteSpace(text))
                dialog.Text = text;
            if (null != owner)
            {
                dialog.ShowDialog(owner);
            }
            else
            {
                dialog.StartPosition = FormStartPosition.CenterScreen;
                dialog.ShowDialog();
            }
            dialog.Dispose();
        }

        /// <summary>
        /// Show modal dialog
        /// </summary>
        /// <param name="selectedObject">object as any</param>
        /// <param name="width">dialog width</param>
        /// <param name="height">dialog height</param>
        /// <param name="text">optional dialog caption</param>
        public static void ShowForm(object selectedObject, int width = 300, int height = 300, string text = null)
        {
            PropertyGridDialog dialog = new PropertyGridDialog(selectedObject);
            if (width >= 300)
                dialog.Width = 300;
            if (height >= 300)
                dialog.Height = 300;
            if (!String.IsNullOrWhiteSpace(text))
                dialog.Text = text;
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.ShowDialog();
            dialog.Dispose();
        }

        #endregion
    }
}