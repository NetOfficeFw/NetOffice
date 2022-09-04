using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NetOffice.OfficeApi.Tools.Dialogs;

namespace NetOffice.OfficeApi.Tools.Contribution
{
    /// <summary>
    /// Dialog extensions for NetOffice applications.
    /// </summary>
    public static class DialogUtilsEx
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
            public DialogUtilsEx.DialogType Type { get; private set; }

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
            public DialogUtilsEx.DialogType Type { get; private set; }

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
            internal NonModalDialogValue(DialogUtilsEx.DialogType type, IEnumerable<KeyValuePair<string, object>> arguments)
            {
                Type = type;
                Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
            }

            #endregion

            #region Properties

            /// <summary>
            /// Shown dialog type
            /// </summary>
            internal DialogUtilsEx.DialogType Type { get; private set; }

            /// <summary>
            /// Arguments dependent on dialog type
            /// </summary>
            internal IEnumerable<KeyValuePair<string, object>> Arguments { get; private set; }

            #endregion
        }

        #endregion

        #region Static members

        private static Dictionary<Form, NonModalDialogValue> _openNonModalDialogs = new Dictionary<Form, NonModalDialogValue>();

        #endregion

        #region Events

        /// <summary>
        /// Occurs before a dialog is shown
        /// </summary>
        public static event DialogShowEventHandler DialogShow;

        /// <summary>
        /// Occurs after a dialog has been closed.
        /// </summary>
        public static event DialogShownEventHandler DialogShown;

        private static void RaiseDialogShown(DialogType type, bool suppressed, bool modal, Result result, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            DialogShownEventArgs args = new DialogShownEventArgs(type, suppressed, modal, result, arguments);
            RaiseDialogShown(args);
        }

        private static void RaiseDialogShown(DialogShownEventArgs arguments)
        {
            var @event = DialogShown;
            if (null != @event)
            {
                @event(null, arguments);
            }
        }

        private static void RaiseDialogShow(DialogType type, bool suppressed, bool modal, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            DialogShowEventArgs args = new DialogShowEventArgs(type, suppressed, modal, arguments);
            RaiseDialogShow(args);
        }

        private static void RaiseDialogShow(DialogShowEventArgs arguments)
        {
            var @event = DialogShow;
            if (null != @event)
            {
                @event(null, arguments);
            }
        }

        #endregion

        #region Properties

        ///// <summary>
        ///// Default dialogs localization settings
        ///// </summary>
        //public DialogLocalizationSettings Localization { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        public static void ShowDiagnostics(this DialogUtils utils, object modalOwner, bool modal, Size size)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

            RaiseDialogShow(DialogType.Diagnostics, isCurrentlySuspended, modal, null);
            if (isCurrentlySuspended)
            {
                RaiseDialogShown(DialogType.Diagnostics, true, modal, Result.No, null);
                return;
            }

            IEnumerable<string> consoleMessages = null;
            if (null != utils.Owner && null != utils.Owner.Owner && null != utils.Owner.Owner.Factory)
            {
                if (!utils.Owner.Owner.Factory.IsInitialized)
                {
#pragma warning disable 612, 618
                    utils.Owner.Owner.Factory.Initialize();
#pragma warning restore 612, 618
                }
                consoleMessages = utils.Owner.Owner.Factory.Console.Messages;
            }

            Dialogs.DiagnosticsDialog dlg = new DiagnosticsDialog(new Informations.DiagnosticPairCollection(utils.Owner), consoleMessages);
            OnCreateToolsDialog(dlg, "DiagnosticsDialog", utils.Layout);

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        public static void ShowDiagnostics(this DialogUtils utils, object modalOwner, bool modal)
        {
            ShowDiagnostics(utils, modalOwner, modal, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        public static void ShowDiagnostics(this DialogUtils utils, bool modal)
        {
            ShowDiagnostics(utils, null, modal, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default diagnostics dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        public static void ShowDiagnostics(this DialogUtils utils)
        {
            ShowDiagnostics(utils, null, false, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        /// <param name="allowDetails">allow user to see exception details</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        public static void ShowError(this DialogUtils utils, object modalOwner, Exception error, string friendlyErrorDescription, bool allowDetails, bool modal, Size size)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
            OnCreateToolsDialog(dlg, "ErrorDialog", utils.Layout);

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        public static void ShowError(this DialogUtils utils, object modalOwner, Exception error, string friendlyErrorDescription)
        {
            ShowError(utils, modalOwner, error, friendlyErrorDescription, true, true, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="error">occured error to display</param>
        /// <param name="friendlyErrorDescription">User-friendly error message to explain what happen</param>
        public static void ShowError(this DialogUtils utils, Exception error, string friendlyErrorDescription)
        {
            ShowError(utils, null, error, friendlyErrorDescription, true, true, Size.Empty);
        }

        /// <summary>
        /// Show the NetOffice default error dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="kind">The method where the error comes from</param>
        /// <param name="error">occured error to display</param>
        public static void ShowErrorDefault(this DialogUtils utils, NetOffice.Tools.ErrorMethodKind kind, Exception error)
        {
            ShowError(utils, null, error, kind.ToString(), true, true, Size.Empty);
        }

        /// <summary>
        /// Show an (un)register error
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">caption</param>
        /// <param name="methodKind">The method where the error comes from</param>
        /// <param name="exception">occured error to display</param>
        public static void ShowRegisterError(this DialogUtils utils, string caption, NetOffice.Tools.RegisterErrorMethodKind methodKind, Exception exception)
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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        public static void ShowRegister(this DialogUtils utils, string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeRegisterKeyState keyState)
        {
            ShowRegister(utils, caption, type, registerCall, scope, keyState, 0);
        }

        /// <summary>
        /// Show message box with register values
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        public static void ShowRegister(this DialogUtils utils, string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeRegisterKeyState keyState, int timeoutSeconds)
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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        public static void ShowUnregister(this DialogUtils utils, string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeUnRegisterKeyState keyState)
        {
            ShowUnregister(utils, caption, type, registerCall, scope, keyState, 0);
        }

        /// <summary>
        /// Show message box with unregister values
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">message box caption</param>
        /// <param name="type">type to register</param>
        /// <param name="registerCall">call kind</param>
        /// <param name="scope">current scope</param>
        /// <param name="keyState">current key state</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        public static void ShowUnregister(this DialogUtils utils, string caption, Type type, NetOffice.Tools.RegisterCall registerCall, NetOffice.Tools.InstallScope scope, NetOffice.Tools.OfficeUnRegisterKeyState keyState, int timeoutSeconds)
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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="assemblyTitle">title of the owner assembly</param>
        /// <param name="assemblyVersion">version of the owner assembly</param>
        /// <param name="copyrightHint">copyright hints of the owner assembly</param>
        /// <param name="companyName">name of the manufactor</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence information</param>
        public static void ShowAbout(this DialogUtils utils, object modalOwner, bool modal, Size size, string headerCaption, string assemblyTitle, string assemblyVersion, string copyrightHint, string companyName, string companyUrl, string licenceText)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
            OnCreateToolsDialog(dlg, "AboutDialog", utils.Layout);

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence information</param>
        public static void ShowAbout(this DialogUtils utils, string headerCaption, string companyUrl, string licenceText)
        {
            ShowAbout(utils, null, true, Size.Empty, headerCaption, utils.Owner.Infos.Assembly.AssemblyTitle, utils.Owner.Infos.Assembly.AssemblyVersion, utils.Owner.Infos.Assembly.AssemblyCopyright, utils.Owner.Infos.Assembly.AssemblyCompany, companyUrl, licenceText);
        }

        /// <summary>
        /// Show the NetOffice default about dialog
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="headerCaption">header caption on top</param>
        /// <param name="companyUrl">optional url of the manufactor</param>
        /// <param name="licenceText">licence information</param>
        public static void ShowAbout(this DialogUtils utils, object modalOwner, bool modal, Size size, string headerCaption, string companyUrl, string licenceText)
        {
            ShowAbout(utils, modalOwner, modal, size, headerCaption, utils.Owner.Infos.Assembly.AssemblyTitle, utils.Owner.Infos.Assembly.AssemblyVersion, utils.Owner.Infos.Assembly.AssemblyCopyright, utils.Owner.Infos.Assembly.AssemblyCompany, companyUrl, licenceText);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
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
        public static Result ShowText(this DialogUtils utils, object modalOwner, string caption, string text, string checkText, bool modal, Size size, int timeoutSeconds, bool skipOnUserAction, Result defaultResult)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
            OnCreateToolsDialog(dlg, "RichTextDialog", utils.Layout);

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public static Result ShowText(this DialogUtils utils, object modalOwner, string caption, string text, string checkText, bool modal, Size size, Result defaultResult)
        {
            return ShowText(utils, modalOwner, caption, text, checkText, modal, size, 0, false, defaultResult);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="size">size for the dialog. Size.Empty to use default size</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public static Result ShowText(this DialogUtils utils, string caption, string text, string checkText, bool modal, Size size, Result defaultResult)
        {
            return ShowText(utils, null, caption, text, checkText, modal, size, defaultResult);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="checkText">additional checkbox want be shown if set. If its true, the checkbox must be checked for result DialogResult.Ok</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public static Result ShowText(this DialogUtils utils, string caption, string text, string checkText, Result defaultResult)
        {
            return ShowText(utils, null, caption, text, checkText, true, Size.Empty, defaultResult);
        }

        /// <summary>
        /// Shows multi-line/rich text to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">text to display. rich text is supported</param>
        /// <param name="timeoutSeconds">timeout in seconds</param>
        /// <param name="skipOnUserAction">skip timeout on user action</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>Result, always none if not modal</returns>
        public static Result ShowText(this DialogUtils utils, string caption, string text, int timeoutSeconds, bool skipOnUserAction, Result defaultResult)
        {
            return ShowText(utils, null, caption, text, null, true, Size.Empty, timeoutSeconds, skipOnUserAction, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, object modalOwner, string text, string caption, Buttons buttons, MessageIcon icon, Result defaultResult)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <param name="icon">icon to show</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, string caption, Buttons buttons, MessageIcon icon, Result defaultResult)
        {
            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, MessageIcon icon, Result defaultResult)
        {
            return ShowMessageBoxInternal(utils, null, text, utils.Owner.Infos.Assembly.AssemblyTitle, Buttons.OK, icon, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, string caption, MessageIcon icon, Result defaultResult)
        {
            return ShowMessageBoxInternal(utils, null, text, caption, Buttons.OK, icon, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, string caption, Buttons buttons, Result defaultResult)
        {
            return ShowMessageBoxInternal(utils, null, text, caption, buttons, MessageIcon.None, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, string caption, Result defaultResult)
        {
            return ShowMessageBoxInternal(utils, null, text, caption, Buttons.OK, MessageIcon.None, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="text">text to display</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        public static Result ShowMessageBox(this DialogUtils utils, string text, Result defaultResult)
        {
            return ShowMessageBoxInternal(utils, null, text, utils.Owner.Infos.Assembly.AssemblyTitle, Buttons.OK, MessageIcon.None, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialogInstance">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="arguments">custom arguments</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public static Result ShowDialog(this DialogUtils utils, object modalOwner, object dialogInstance, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, Result defaultResult)
        {
            Form dialog = (Form)dialogInstance;

            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public static Result ShowDialog(this DialogUtils utils, object modalOwner, object dialog, bool modal)
        {
            return ShowDialog(utils, modalOwner, (Form)dialog, modal, null, Result.None);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public static Result ShowDialog(this DialogUtils utils, object modalOwner, object dialog, bool modal, Result defaultResult)
        {
            return ShowDialog(utils, modalOwner, (Form)dialog, modal, null, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="arguments">custom arguments</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public static Result ShowDialog(this DialogUtils utils, object dialog, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, Result defaultResult)
        {
            return ShowDialog(utils, null, (Form)dialog, modal, arguments, defaultResult);
        }

        /// <summary>
        /// Show dialog instance
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="dialog">dialog instance to show</param>
        /// <param name="modal">show dialog modal to its owner window</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>DialogResult, always none if not modal</returns>
        public static Result ShowDialog(this DialogUtils utils, object dialog, bool modal, Result defaultResult)
        {
            return ShowDialog(utils, null, (Form)dialog, modal, null, defaultResult);
        }

        /// <summary>
        /// Show modal Windows.Forms message box to the user
        /// </summary>
        /// <param name="utils">DialogUtils instance</param>
        /// <param name="modalOwner">owner window. can be null(Nothing in Visual Basic)</param>
        /// <param name="text">text to display</param>
        /// <param name="caption">dialog title</param>
        /// <param name="buttons">user selection buttons</param>
        /// <param name="icon">default icon</param>
        /// <param name="defaultResult">result if its not shown</param>
        /// <returns>user selection</returns>
        internal static Result ShowMessageBoxInternal(this DialogUtils utils, object modalOwner, string text, string caption, Buttons buttons, MessageIcon icon, Result defaultResult)
        {
            IWin32Window owner = Running.Win32Window.Create(modalOwner);

            bool isCurrentlySuspended = utils.IsCurrentlySuspended();

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
        /// Called after create a new ToolsDialog instance
        /// </summary>
        /// <param name="dialog">new instance</param>
        /// <param name="dialogName">name of the dialog</param>
        /// <param name="layout">Dialog layout settings</param>
        private static void OnCreateToolsDialog(ToolsDialog dialog, string dialogName, DialogLayoutSettings layout)
        {
            //dialog.DoLocalization(Localization[dialogName][CurrentLanguage, true]);
            dialog.DoLayout(layout);
        }

        /// <summary>
        /// Displays a message box with specified text
        /// </summary>
        /// <param name="text">specified text</param>
        public static void ShowMessageBox(string text)
        {
            MessageBox.Show(text);
        }

        /// <summary>
        /// Displays a message box with specified text and caption
        /// </summary>
        /// <param name="text">specified text</param>
        /// <param name="caption">The text display in the title bar</param>
        public static void ShowMessageBox(string text, string caption)
        {
            MessageBox.Show(text, caption);
        }

        #endregion

        #region Trigger

        private static void NonModalDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Form formSender = sender as Form;
                if (formSender == null)
                {
                    return;
                }

                formSender.FormClosed -= NonModalDialog_FormClosed;
                if (_openNonModalDialogs.ContainsKey(formSender))
                {
                    NonModalDialogValue formValue = _openNonModalDialogs[formSender];
                    RaiseDialogShown(formValue.Type, false, false, (Result)formSender.DialogResult, formValue.Arguments);
                }
            }
            catch (Exception exception)
            {
                Core.Default.Console.WriteException(exception);
            }
        }

        #endregion
    }
}