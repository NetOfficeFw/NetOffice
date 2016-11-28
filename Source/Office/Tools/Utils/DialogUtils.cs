using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Text;
using NetOffice.OfficeApi.Tools.Dialogs;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Dialog related utils
    /// </summary>
    public class DialogUtils
    {
        #region Embedded Definitions

        /// <summary>
        /// Specifies constants defining which information to display.
        /// </summary>
        public enum MessageIcon
        {
            /// <summary>
            /// The message box contain no symbols.
            /// </summary>
            None = 0,

            /// <summary>
            /// The message box contains a symbol consisting of a white X in a circle with a  red background.
            /// </summary>
            Hand = 16,

            /// <summary>
            /// The message box contains a symbol consisting of white X in a circle with a red background.
            /// </summary>
            Stop = 16,

            /// <summary>
            /// The message box contains a symbol consisting of white X in a circle with a red background.
            /// </summary>
            Error = 16,

            /// <summary>
            /// The message box contains a symbol consisting of a question mark in a circle.
            /// </summary>
            Question = 32,

            /// <summary>
            /// The message box contains a symbol consisting of an exclamation point in a triangle with a yellow background.
            /// </summary>
            Exclamation = 48,

            /// <summary>
            /// The message box contains a symbol consisting of an exclamation point in a triangle with a yellow background.
            /// </summary>
            Warning = 48,

            /// <summary>
            /// The message box contains a symbol consisting of a lowercase letter i in a circle.
            /// </summary>
            Asterisk = 64,

            /// <summary>
            /// The message box contains a symbol consisting of a lowercase letter i in a circle.
            /// </summary>
            Information = 64
        }

        /// <summary>
        /// Specifies constants defining which buttons to display
        /// </summary>
        public enum Buttons
        {
            /// <summary>
            /// The message box contains an OK button.
            /// </summary>
            OK = 0,

            /// <summary>
            /// The message box contains OK and Cancel buttons.
            /// </summary>
            OKCancel = 1,

            /// <summary>
            /// The message box contains Abort, Retry, and Ignore buttons.
            /// </summary>
            AbortRetryIgnore = 2,

            /// <summary>
            /// The message box contains Yes, No, and Cancel buttons.
            /// </summary>
            YesNoCancel = 3,

            /// <summary>
            /// The message box contains Yes and No buttons.
            /// </summary>
            YesNo = 4,

            /// <summary>
            /// The message box contains Retry and Cancel buttons.
            /// </summary>
            RetryCancel = 5
        }

        /// <summary>
        /// Specifies identifiers to indicate the return value of a dialog box.
        /// </summary>
        public enum Result
        {
            /// <summary>
            /// Nothing is returned from the dialog box. This means that the modal dialog continues running.
            /// </summary>
            None = 0,

            /// <summary>
            /// The dialog box return value is OK (usually sent from a button labeled OK).
            /// </summary>
            OK = 1,

            /// <summary>
            /// The dialog box return value is Cancel (usually sent from a button labeled Cancel).
            /// </summary>
            Cancel = 2,

            /// <summary>
            /// The dialog box return value is Abort (usually sent from a button labeled Abort).
            /// </summary>
            Abort = 3,

            /// <summary>
            /// The dialog box return value is Retry (usually sent from a button labeled Retry).
            /// </summary>
            Retry = 4,

            /// <summary>
            /// The dialog box return value is Ignore (usually sent from a button labeled Ignore).
            /// </summary>
            Ignore = 5,

            /// <summary>
            /// The dialog box return value is Yes (usually sent from a button labeled Yes).
            /// </summary>
            Yes = 6,

            /// <summary>
            ///  The dialog box return value is No (usually sent from a button labeled No).
            /// </summary>
            No = 7
        }

        /// <summary>
        /// Indicates which kind of dialog is shown
        /// </summary>
        public enum DialogType
        {
            /// <summary>
            /// Custom dialog instance
            /// </summary>
            Custom = 0,

            /// <summary>
            /// Windows.Forms MessageBox
            /// </summary>
            MessageBox = 1,

            /// <summary>
            /// Error Dialog
            /// </summary>
            Error = 2,

            /// <summary>
            /// About Dialog
            /// </summary>
            About = 3,

            /// <summary>
            /// Diagnostics Dialog
            /// </summary>
            Diagnostics = 4,

            /// <summary>
            /// Multi-Line Text Dialog, also RichText is supported
            /// </summary>
            Text = 5
        }

        /// <summary>
        /// Dialog show event arguments
        /// </summary>
        public class DialogShowEventArgs : EventArgs
        {
            #region Ctor

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="type">dialog type</param>
            /// <param name="suppressed">dialog want not shown</param>
            /// <param name="modal">dialog want shown as modal to its parent</param>
            /// <param name="arguments">arguments dependent on dialog type</param>
            internal DialogShowEventArgs(DialogType type, bool suppressed, bool modal, IEnumerable<KeyValuePair<string, object>> arguments)
            {
                Type = type;
                Suppressed = suppressed;
                Modal = modal;
                Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
            }

            #endregion

            #region Properties

            /// <summary>
            /// Dialog want shown as modal to its parent.
            /// </summary>
            public bool Modal { get; private set; }

            /// <summary>
            /// The dialog want not shown because its currently forbidden by dialog settings
            /// </summary>
            public bool Suppressed { get; private set; }

            /// <summary>
            /// Shown dialog type
            /// </summary>
            public DialogUtils.DialogType Type { get; private set; }

            /// <summary>
            /// Arguments dependent on dialog type
            /// </summary>
            public IEnumerable<KeyValuePair<string, object>> Arguments { get; private set; }

            #endregion
        }

        /// <summary>
        /// Dialog shown event arguments
        /// </summary>
        public class DialogShownEventArgs : EventArgs
        {
            #region Ctor

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="type">dialog type</param>
            /// <param name="suppressed">dialog has not shown</param>
            /// <param name="modal">dialog has shown as modal to its parent</param>
            /// <param name="result">dialog result if set</param>
            /// <param name="arguments">arguments dependent on dialog type</param>
            internal DialogShownEventArgs(DialogType type, bool suppressed, bool modal, Result result, IEnumerable<KeyValuePair<string, object>> arguments)
            {
                Type = type;
                Suppressed = suppressed;
                Modal = modal;
                Result = result;
                Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
            }

            #endregion

            #region Properties

            /// <summary>
            /// Dialog has shown as modal to its parent.
            /// </summary>
            public bool Modal { get; private set; }

            /// <summary>
            /// The dialog has not shown because its currently forbidden by dialog settings
            /// </summary>
            public bool Suppressed { get; private set; }

            /// <summary>
            /// Dialog result if set
            /// </summary>
            public Result Result { get; private set; }

            /// <summary>
            /// Shown dialog type
            /// </summary>
            public DialogUtils.DialogType Type { get; private set; }

            /// <summary>
            /// Arguments dependent on dialog type
            /// </summary>
            public IEnumerable<KeyValuePair<string, object>> Arguments { get; private set; }

            #endregion
        }

        /// <summary>
        /// Dialog shown event handler
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="arguments">dialog shown arguments</param>
        public delegate void DialogShownEventHandler(DialogUtils sender, DialogShownEventArgs arguments);

        /// <summary>
        /// Dialog show event handler
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="arguments">dialog show arguments</param>
        public delegate void DialogShowEventHandler(DialogUtils sender, DialogShowEventArgs arguments);

        /// <summary>
        /// Encapsulate caller arguments to observe non-modal shown dialogs and fire DialogShown event after close
        /// </summary>
        private class NonModalDialogValue
        {
            #region Ctor

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="type">shown dialog type</param>
            /// <param name="arguments">arguments dependent on dialog type</param>
            internal NonModalDialogValue(DialogUtils.DialogType type, IEnumerable<KeyValuePair<string, object>> arguments)
            {
                Type = type;
                Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
            }

            #endregion

            #region Properties

            /// <summary>
            /// Shown dialog type
            /// </summary>
            internal DialogUtils.DialogType Type { get; private set; }

            /// <summary>
            /// Arguments dependent on dialog type
            /// </summary>
            internal IEnumerable<KeyValuePair<string, object>> Arguments { get; private set; }

            #endregion
        }

        #endregion

        #region Fields

        private const int _currentDefaultLanguage = 1033;
        private CommonUtils _owner;
        private Dictionary<Form, NonModalDialogValue> _openNonModalDialogs;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal DialogUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            CurrentLanguage = _currentDefaultLanguage;
            _owner = owner;
            _openNonModalDialogs = new Dictionary<Form, NonModalDialogValue>();
            SuppressOnAutomation = true;
            SuppressOnHide = true;
            Layout = new DialogLayoutSettings();
            Localization = new DialogLocalizationSettings(ToolsDialog.CreateDialogSchema());
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs before a dialog is shown
        /// </summary>
        public event DialogShowEventHandler DialogShow;

        /// <summary>
        /// Occurs after a dialog has been closed. 
        /// </summary>
        public event DialogShownEventHandler DialogShown;

        private void RaiseDialogShown(DialogType type, bool suppressed, bool modal, Result result, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            DialogShownEventArgs args = new DialogShownEventArgs(type, suppressed, modal, result, arguments);
            RaiseDialogShown(args);
        }

        private void RaiseDialogShown(DialogShownEventArgs arguments)
        {
            if (null != DialogShown)
                DialogShown(this, arguments);
        }

        private void RaiseDialogShow(DialogType type, bool suppressed, bool modal, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            DialogShowEventArgs args = new DialogShowEventArgs(type, suppressed, modal, arguments);
            RaiseDialogShow(args);
        }

        private void RaiseDialogShow(DialogShowEventArgs arguments)
        {
            if (null != DialogShow)
                DialogShow(this, arguments);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Current used language in dialogs. Default is 1033(en-us) If its failed to find a dialog localization set for current language en-us want be used
        /// </summary>
        public int CurrentLanguage { get; set; }

        /// <summary>
        /// Dont show dialogs if office application is started programmatically for automation 
        /// </summary>
        public bool SuppressOnAutomation { get; set; }

        /// <summary>
        /// Dont show dialogs if office application is currently not visible
        /// </summary>
        public bool SuppressOnHide { get; set; }

        /// <summary>
        /// Dont show dialogs at all
        /// </summary>
        public bool SupressGeneraly { get; set; }

        /// <summary>
        /// Default dialogs layout settings
        /// </summary>
        public DialogLayoutSettings Layout { get; private set; }

        /// <summary>
        /// Default dialogs localization settings
        /// </summary>
        public DialogLocalizationSettings Localization { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Returns information show dialogs is currently suspended
        /// </summary>
        /// <returns>true if suspended otherwise false</returns>
        public virtual bool IsCurrentlySuspended()
        {
            if (SupressGeneraly)
                return true;
            if (SuppressOnAutomation && _owner.IsAutomation)
                return true;
            if (SuppressOnHide && false == TryGetApplicationVisible(true))
                return true;
            return false;
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        public virtual void ShowDiagnostics(object modalOwner, bool modal, Size size)
        {            
            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            RaiseDialogShow(DialogType.Diagnostics, isCurrentlySuspended, modal, null);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.Diagnostics, true, modal, Result.No, null);
                return;
            }

            IEnumerable<string> consoleMessages = null;
            if (null != _owner && null != _owner.Owner && null != _owner.Owner.Factory)
            {
                if (!_owner.Owner.Factory.IsInitialized)
                {
#pragma warning disable 612, 618
                    _owner.Owner.Factory.Initialize();
#pragma warning restore 612, 618
                }
                consoleMessages = _owner.Owner.Factory.Console.Messages;
            }

            Dialogs.DiagnosticsDialog dlg = new DiagnosticsDialog(new Informations.DiagnosticPairCollection(_owner), consoleMessages);
            OnCreateToolsDialog(dlg, "DiagnosticsDialog");

            if (null == owner)
                dlg.StartPosition = FormStartPosition.CenterScreen;
            if (Size.Empty != size)
                dlg.Size = size;

            if (modal)
            {
                dlg.ShowDialog(owner);
                RaiseDialogShown(DialogType.Diagnostics, false, true, Result.No, null);
            }
            else
            {
                _openNonModalDialogs.Add(dlg, new NonModalDialogValue(DialogType.Diagnostics, null));
                dlg.FormClosed += new FormClosedEventHandler(NonModalDialog_FormClosed);
                dlg.Show(owner);
            }
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        public void ShowDiagnostics(object modalOwner, bool modal)
        {
            ShowDiagnostics(modalOwner, modal, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="modal">show dialog modal to its owner window</param>
        public void ShowDiagnostics(bool modal)
        {
            ShowDiagnostics(null, modal, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        public void ShowDiagnostics()
        {
            ShowDiagnostics(null, false, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        /// <param name="allowDetails">allow user to see exception details</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        public virtual void ShowError(object modalOwner, Exception error, string friendlyErrorDescription, bool allowDetails, bool modal, Size size)
        {
            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            List<KeyValuePair<string, object>> arguments = new List<KeyValuePair<string, object>>();
            arguments.Add(new KeyValuePair<string, object>("Error", error));
            arguments.Add(new KeyValuePair<string, object>("Description", friendlyErrorDescription));

            RaiseDialogShow(DialogType.Error, isCurrentlySuspended, modal, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.Error, true, modal, Result.No, arguments);
                return;
            }

            Dialogs.ErrorDialog dlg = new ErrorDialog(error, friendlyErrorDescription, allowDetails);
            OnCreateToolsDialog(dlg, "ErrorDialog");

            if (null == owner)
                dlg.StartPosition = FormStartPosition.CenterScreen;
            if (Size.Empty != size)
                dlg.Size = size;

            if (modal)
            {
                dlg.ShowDialog(owner);
                RaiseDialogShown(DialogType.Error, false, true, Result.No, arguments);
            }
            else
            {
                _openNonModalDialogs.Add(dlg, new NonModalDialogValue(DialogType.Error, arguments));
                dlg.FormClosed += new FormClosedEventHandler(NonModalDialog_FormClosed);
                dlg.Show(owner);
            }
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        public void ShowError(object modalOwner, Exception error, string friendlyErrorDescription)
        {
            ShowError(modalOwner, error, friendlyErrorDescription, true, true, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        public void ShowError(Exception error, string friendlyErrorDescription)
        {
            ShowError(null, error, friendlyErrorDescription, true, true, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="kind">The method where the error comes from</param>
        /// <param name="error">occured error to display</param>
        public void ShowErrorDefault(NetOffice.Tools.ErrorMethodKind kind, Exception error)
        {
            ShowError(null, error, kind.ToString(), true, true, Size.Empty);
        }

        /// <summary>
        /// Show an (un)register error
        /// </summary>
        /// <param name="caption">caption</param>
        /// <param name="methodKind">The method where the error comes from</param>
        /// <param name="exception">occured error to display</param>
        public static void ShowRegisterError(string caption, NetOffice.Tools.RegisterErrorMethodKind methodKind, Exception exception)
        {
            if (null == caption || "" == caption)
                caption = methodKind.ToString() + "  Error";
            
            string text = methodKind.ToString() + "  Error" + Environment.NewLine + Environment.NewLine;
            if (null != exception)
                text += exception.ToString();

            MessageBox.Show(text, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Show message box with register values
        /// </summary>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        public static void ShowRegister(string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeRegisterKeyState keyState)
        {
            ShowRegister(caption, type, registerCall, scope, keyState, 0);
        }

        /// <summary>
        /// Show message box with register values
        /// </summary>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        public static void ShowRegister(string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeRegisterKeyState keyState, int timeoutSeconds)
        {
            string text = String.Format("Type: {0}{4}RegisterCall: {1}{4}Scope:{2}{4}KeyState: {3}{4}",
                null != type ? type.ToString() : "<Empty>", registerCall, scope, keyState, Environment.NewLine);
            RichTextDialog dlg = new RichTextDialog("Register " + caption, text, timeoutSeconds, true);
            dlg.Text = "Register";
            dlg.TopMost = true;
            dlg.ShowInTaskbar = true;
            dlg.StartPosition = FormStartPosition.CenterScreen;
            dlg.ShowDialog();
        }

        /// <summary>
        /// Show message box with unregister values
        /// </summary>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        public static void ShowUnregister(string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeUnRegisterKeyState keyState)
        {
            ShowUnregister(caption, type, registerCall, scope, keyState, 0);
        }

        /// <summary>
        /// Show message box with unregister values
        /// </summary>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        public static void ShowUnregister(string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeUnRegisterKeyState keyState, int timeoutSeconds)
        {
            string text = String.Format("Type: {0}{4}RegisterCall: {1}{4}Scope: {2}{4}KeyState: {3}{4}",
                null != type ? type.ToString() : "<Empty>", registerCall, scope, keyState, Environment.NewLine);

            RichTextDialog dlg = new RichTextDialog("Unregister " + caption, text, timeoutSeconds, true);
            dlg.Text = "Unregister";
            dlg.TopMost = true;
            dlg.ShowInTaskbar = true;
            dlg.StartPosition = FormStartPosition.CenterScreen;
            dlg.ShowDialog();
        }

        /// <summary>
        /// Show the NetOffice default about dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="assemblyTitle">title of the owner assembly</param>
        /// <param name="assemblyVersion">version of the owner assembly</param>
        /// <param name="copyrightHint">copyright hints of the owner assembly</param>
        /// <param name="companyName">name of the manufactor</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence informations</param>
        public void ShowAbout(object modalOwner, bool modal, Size size, string headerCaption, string assemblyTitle, string assemblyVersion, string copyrightHint, string companyName, string companyUrl, string licenceText)
        {
            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            List<KeyValuePair<string, object>> arguments = new List<KeyValuePair<string, object>>();
            arguments.Add(new KeyValuePair<string, object>("Caption", headerCaption));
            arguments.Add(new KeyValuePair<string, object>("AssemblyTitle", assemblyTitle));
            arguments.Add(new KeyValuePair<string, object>("AssemblyVersion", assemblyVersion));
            arguments.Add(new KeyValuePair<string, object>("Copyright", copyrightHint));
            arguments.Add(new KeyValuePair<string, object>("CompanyName", companyName));
            arguments.Add(new KeyValuePair<string, object>("CompanyUrl", companyUrl));
            arguments.Add(new KeyValuePair<string, object>("Licence", licenceText));

            RaiseDialogShow(DialogType.About, isCurrentlySuspended, modal, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.About, true, modal, Result.None, arguments);
                return;
            }

            Dialogs.AboutDialog dlg = new AboutDialog(headerCaption, assemblyTitle, assemblyVersion, copyrightHint, companyName, companyUrl, licenceText);
            OnCreateToolsDialog(dlg, "AboutDialog");

            if (null == owner)
                dlg.StartPosition = FormStartPosition.CenterScreen;
            if (Size.Empty != size)
                dlg.Size = size;

            if (modal)
            {
                dlg.ShowDialog(owner);
                RaiseDialogShown(DialogType.About, false, true, Result.None, arguments);
            }
            else
            {
                _openNonModalDialogs.Add(dlg, new NonModalDialogValue(DialogType.About, arguments));
                dlg.FormClosed += new FormClosedEventHandler(NonModalDialog_FormClosed);
                dlg.Show(owner);
            }
        }

        /// <summary>
        /// Show the NetOffice default about dialog
        /// </summary>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence informations</param>
        public void ShowAbout(string headerCaption, string companyUrl, string licenceText)
        {
            ShowAbout(null, true, Size.Empty, headerCaption, _owner.Infos.Assembly.AssemblyTitle, _owner.Infos.Assembly.AssemblyVersion, _owner.Infos.Assembly.AssemblyCopyright, _owner.Infos.Assembly.AssemblyCompany, companyUrl, licenceText);
        }

        /// <summary>
        /// Show the NetOffice default about dialog
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence informations</param>
        public void ShowAbout(object modalOwner, bool modal, Size size, string headerCaption, string companyUrl, string licenceText)
        {
            ShowAbout(modalOwner, modal, size, headerCaption, _owner.Infos.Assembly.AssemblyTitle, _owner.Infos.Assembly.AssemblyVersion, _owner.Infos.Assembly.AssemblyCopyright, _owner.Infos.Assembly.AssemblyCompany, companyUrl, licenceText);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        /// <param name="skipOnUserAction">skip timeout on user action</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public virtual Result ShowText(object modalOwner, string caption, string text, string checkText, bool modal, Size size, int timeoutSeconds, bool skipOnUserAction, Result defaultResult)
        {
            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            List<KeyValuePair<string, object>> arguments = new List<KeyValuePair<string, object>>();
            arguments.Add(new KeyValuePair<string, object>("Caption", caption));
            arguments.Add(new KeyValuePair<string, object>("Text", text));
            arguments.Add(new KeyValuePair<string, object>("CheckText", checkText));

            RaiseDialogShow(DialogType.Text, isCurrentlySuspended, modal, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.Text, true, modal, defaultResult, arguments);
                return defaultResult;
            }

            Dialogs.RichTextDialog dlg = new RichTextDialog(caption, text, checkText, timeoutSeconds, skipOnUserAction);
            OnCreateToolsDialog(dlg, "RichTextDialog");

            if (null == owner)
                dlg.StartPosition = FormStartPosition.CenterScreen;
            if (Size.Empty != size)
                dlg.Size = size;

            if (modal)
            {
                DialogResult dlgResult = dlg.ShowDialog(owner);
                RaiseDialogShown(DialogType.Text, false, true, (Result)dlgResult, arguments);
                return (Result)dlgResult;
            }
            else
            {
                _openNonModalDialogs.Add(dlg, new NonModalDialogValue(DialogType.Text, arguments));
                dlg.FormClosed += new FormClosedEventHandler(NonModalDialog_FormClosed);
                dlg.Show(owner);
                return Result.None;
            }
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public virtual Result ShowText(object modalOwner, string caption, string text, string checkText, bool modal, Size size, Result defaultResult)
        {
            return ShowText(modalOwner, caption, text, checkText, modal, size, 0, false, defaultResult);
        
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public virtual Result ShowText(string caption, string text, string checkText, bool modal, Size size, Result defaultResult)
        {
            return ShowText(null, caption, text, checkText, modal, size, defaultResult);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public virtual Result ShowText(string caption, string text, string checkText, Result defaultResult)
        {
            return ShowText(null, caption, text, checkText, true, Size.Empty, defaultResult);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        /// <param name="skipOnUserAction">skip timeout on user action</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public virtual Result ShowText(string caption, string text, int timeoutSeconds, bool skipOnUserAction, Result defaultResult)
        {
            return ShowText(null, caption, text, null, true, Size.Empty, timeoutSeconds, skipOnUserAction, defaultResult);
        }
        
        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="arguments">given arguments as any to use like String.Format in text</param>
        /// <returns>Result.OK</returns>
        public Result ShowMessageBox(string text, params object[] arguments)
        {
            string validatedText = String.Format(text, arguments);
            return ShowMessageBox(null, validatedText, null, MessageBoxButtons.OK, MessageBoxIcon.None, DialogResult.OK);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(object modalOwner, string text, string caption, Buttons buttons, MessageIcon icon, Result defaultResult)
        {
            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            List<KeyValuePair<string, object>> arguments = new List<KeyValuePair<string, object>>();
            arguments.Add(new KeyValuePair<string, object>("Caption", caption));
            arguments.Add(new KeyValuePair<string, object>("Text", text));

            RaiseDialogShow(DialogType.MessageBox, isCurrentlySuspended, true, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.MessageBox, true, true, defaultResult, arguments);
                return defaultResult;
            }

            DialogResult dlgResult = MessageBox.Show(owner, text, caption, (MessageBoxButtons)buttons, (MessageBoxIcon)icon);
            RaiseDialogShown(DialogType.MessageBox, false, true, (Result)dlgResult, arguments);
            return (Result)dlgResult;
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <param name="icon">icon to show</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, string caption, Buttons buttons, MessageIcon icon, Result defaultResult)
        {           
            bool isCurrentlySuspended = IsCurrentlySuspended();

            List<KeyValuePair<string, object>> arguments = new List<KeyValuePair<string, object>>();
            arguments.Add(new KeyValuePair<string, object>("Caption", caption));
            arguments.Add(new KeyValuePair<string, object>("Text", text));

            RaiseDialogShow(DialogType.MessageBox, isCurrentlySuspended, true, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.MessageBox, true, true, defaultResult, arguments);
                return defaultResult;
            }

            DialogResult dlgResult = MessageBox.Show(null, text, caption, (MessageBoxButtons)buttons, (MessageBoxIcon)icon);
            RaiseDialogShown(DialogType.MessageBox, false, true, (Result)dlgResult, arguments);
            return (Result)dlgResult;
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, MessageIcon icon, Result defaultResult)
        {
            return ShowMessageBox(null, text, _owner.Infos.Assembly.AssemblyTitle, MessageBoxButtons.OK, icon, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, string caption, MessageIcon icon, Result defaultResult)
        {
            return ShowMessageBox(null, text, caption, MessageBoxButtons.OK, icon, defaultResult);   
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, string caption, Buttons buttons, Result defaultResult)
        {
            return ShowMessageBox(null, text, caption, buttons, MessageBoxIcon.None, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, string caption, Result defaultResult)
        {
            return ShowMessageBox(null, text, caption, MessageBoxButtons.OK, MessageBoxIcon.None, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="text">text to display</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public Result ShowMessageBox(string text, Result defaultResult)
        {
            return ShowMessageBox(null, text, _owner.Infos.Assembly.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.None, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialogInstance">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="arguments">custom arguments</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public Result ShowDialog(object modalOwner, object dialogInstance, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, Result defaultResult)
        {
            Form dialog = (Form)dialogInstance;

            IWin32Window owner = NetOffice.Tools.WndUtils.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = IsCurrentlySuspended();

            if(null == arguments)
                 arguments = new List<KeyValuePair<string, object>>();

            RaiseDialogShow(DialogType.Custom, isCurrentlySuspended, modal, arguments);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.Custom, true, modal, defaultResult, arguments);
                return defaultResult;
            }

            if (modal)
            {
                DialogResult dlgResult = dialog.ShowDialog(owner);
                RaiseDialogShown(DialogType.Custom, false, true, (Result)dlgResult, arguments);
                return (Result)dlgResult;
            }
            else
            {
                _openNonModalDialogs.Add(dialog, new NonModalDialogValue(DialogType.Custom, arguments));
                dialog.FormClosed += new FormClosedEventHandler(NonModalDialog_FormClosed);
                dialog.Show(owner);
                return Result.None;
            }
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public Result ShowDialog(object modalOwner, object dialog, bool modal)
        {
            return ShowDialog(modalOwner, (Form)dialog, modal, null, Result.None);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public Result ShowDialog(object modalOwner, object dialog, bool modal, Result defaultResult)
        {
            return ShowDialog(modalOwner, (Form)dialog, modal, null, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="arguments">custom arguments</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public Result ShowDialog(object dialog, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, Result defaultResult)
        {
            return ShowDialog(null, (Form)dialog, modal, arguments, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public Result ShowDialog(object dialog, bool modal, Result defaultResult)
        {
            return ShowDialog(null, (Form)dialog, modal, null, defaultResult);
        }

        /// <summary>
        /// Called if its failed to proceed a non-modal dialog after close
        /// </summary>
        /// <param name="exception">unexpected origin error</param>
        protected internal virtual void OnDialogError(Exception exception)
        { 
        
        }

        /// <summary>
        /// Called after create a new ToolsDialog instance
        /// </summary>
        /// <param name="dialog">new instance</param>
        /// <param name="dialogName">name of the dialog</param>
        protected virtual void OnCreateToolsDialog(ToolsDialog dialog, string dialogName)
        {
            dialog.DoLocalization(Localization[dialogName][CurrentLanguage, true]);
            dialog.DoLayout(Layout);
        }
        
        /// <summary>
        /// Try to detect the visibilty of host application main window.
        /// The implementation want find a Visible property and analyze its current state
        /// </summary>
        /// <param name="defaultResult">fallback result if its failed</param>
        /// <returns>true if application is visible, otherwise false</returns>
        protected virtual bool TryGetApplicationVisible(bool defaultResult)
        {
            try
            {
                if (_owner.OwnerApplication.EntityIsAvailable("Visible"))
                {
                    object result = _owner.OwnerApplication.Invoker.PropertyGet(_owner.OwnerApplication, "Visible");
                    if (result is bool)
                        return Convert.ToBoolean(result);
                    else
                    {
                        int i = Convert.ToInt32(result);
                        return i != 0;
                    }
                }
                else
                {
                    return defaultResult;
                }

            }
            catch (Exception exception)
            {
                OnDialogError(exception);
                return defaultResult;
            }
        }

        #endregion

        #region Trigger

        private void NonModalDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Form formSender = sender as Form;
                if (_openNonModalDialogs.ContainsKey(formSender))
                {
                    NonModalDialogValue formValue = _openNonModalDialogs[formSender];
                    RaiseDialogShown(formValue.Type, false, false, (Result)formSender.DialogResult, formValue.Arguments);
                }
            }
            catch (Exception exception)
            {
                OnDialogError(exception);
            }
        }

        #endregion
    }
}
