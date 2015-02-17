using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Application Language Edit Form
    /// </summary>
    public partial class ToolLanguageForm : Form
    {
        #region API

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        #endregion

        #region Fields

        private Control _highlightControl1;
        private Control _highlightControl2;
        private Pen _highlightPen;        
        private int _selectedTabIndex;
        private ToolLanguage _language;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ToolLanguageForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="language">edit language</param>
        internal ToolLanguageForm(ToolLanguage language)
        {
            InitializeComponent();
            _language = language;
            if (language is ToolDefaultLanguage)
            {
                label1DefaultHint.Text = String.Format("{0} is a readonly default language.", language.NameGlobal);
                panelDefaultHint.Visible = true;
            }
            else
                panelDefaultHint.Visible = false;

            toolLanguageControl1.SelectedLanguage = language;
            _highlightPen = new Pen(Color.Red, 2);
            overlayPainter1.Owner = this;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Unsaved changes on the edit language
        /// </summary>
        public bool Changed { get; private set; }

        /// <summary>
        /// The window is currently visible and in front
        /// </summary>
        private bool IsActive
        {
            get
            {
                return GetForegroundWindow() == this.Handle && this.WindowState != FormWindowState.Minimized;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates an instance of ToolLanguageForm and show
        /// </summary>
        /// <param name="owner">modal owner</param>
        /// <param name="language">edit language</param>
        /// <returns>contains unsaved changes</returns>
        internal static bool ShowForm(IWin32Window owner, ToolLanguage language)
        {
            ToolLanguageForm dlg = new ToolLanguageForm(language);
            dlg.ShowDialog(owner);
            dlg.Dispose(true);
            return dlg.Changed;
        }

        /// <summary>
        /// Draw rectangle arround given control
        /// </summary>
        /// <param name="ctrl">control to highlight</param>
        internal void StartHighLightControl1(Control ctrl)
        {
            StopHighLightControl1();
            _highlightControl1 = ctrl;
        }

        /// <summary>
        /// Stop drawing rectangle arround controls
        /// </summary>
        internal void StopHighLightControl1()
        {
            _highlightControl1 = null;
        }

        /// <summary>
        /// Draw rectangle arround given control
        /// </summary>
        /// <param name="ctrl">control to highlight</param>
        internal void StartHighLightControl2(Control ctrl)
        {
            StopHighLightControl2();
            _highlightControl2 = ctrl;
        }

        /// <summary>
        /// Stop drawing rectangle arround controls
        /// </summary>
        internal void StopHighLightControl2()
        {
            _highlightControl2 = null;
        }

        private Rectangle FindRect(Control ctrl)
        {
            Point controlLoc = ctrl.PointToScreen(Point.Empty);
            Point formLoc = this.PointToScreen(Point.Empty);
            
            Point relativeLoc = new Point(controlLoc.X - formLoc.X, controlLoc.Y - formLoc.Y);
            return new Rectangle(relativeLoc.X, relativeLoc.Y, ctrl.Width, ctrl.Height+1);
        }

        #endregion

        #region Trigger

        private void overlayPainter1_Paint(object sender, PaintEventArgs e)
        {
            if (_selectedTabIndex == 0 || false == IsActive)
                return;

            Control targetControl = _highlightControl1;
            if (_selectedTabIndex ==1 &&  null != targetControl)
            {
                Rectangle rect = FindRect(targetControl);
                e.Graphics.DrawRectangle(_highlightPen, rect);
            }

            targetControl = _highlightControl2;
            if (_selectedTabIndex == 2 && null != targetControl)
            {
                Rectangle rect = FindRect(targetControl);
                e.Graphics.DrawRectangle(_highlightPen, rect);
            }
        }

        private void toolLanguageControl1_SelectedTabChanged(object sender, EventArgs e)
        {
            _selectedTabIndex = toolLanguageControl1.SelectedTabIndex;
        }

        private void ToolLanguageForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt && e.KeyCode == Keys.Up)
            {
                toolLanguageControl1.HandleKeyUp();
            }
            else if (e.Alt && e.KeyCode == Keys.Down)
            {
                toolLanguageControl1.HandleKeyDown();            
            }
        }

        private void buttonSaveChanges_Click(object sender, EventArgs e)
        {
            if (_language is ToolDefaultLanguage)
            {
                this.Close();
                return;
            }

            if (!_language.IsValid())
            {
                if (DialogResult.No == MessageBox.Show(this, String.Format("Unable to save changes because no global name and/or valid LCID is set.{0}{0}Close anyway?", Environment.NewLine), "Sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                    return;
                else
                {
                    this.Close();
                    return;
                }
            }

            if (_language.IsNew || _language.IsDirty && _language.IsValid())
            {
                try
                {
                    _language.Save();
                    Changed = true;
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                }
            }

            this.Close();
        }

        private void toolLanguageControl1_SelectedNodeTextChanged(object sender, EventArgs e)
        {
            Text = " Edit Language " + toolLanguageControl1.SelectedNodeText;
        }

        #endregion
    }
}
