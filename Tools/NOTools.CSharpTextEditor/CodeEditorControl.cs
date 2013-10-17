using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.CodeDom.Compiler;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Editing;
using NOTools.CSharpTextEditor.GACManagedAccess;

namespace NOTools.CSharpTextEditor
{
    /// <summary>
    /// WindowsForm Wrapper Control for AvalonEdit
    /// </summary>
    public partial class CodeEditorControl : UserControl
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public CodeEditorControl()
        {
            InitializeComponent();
            ErrorPanelSettings = new ErrorPanelOptions(this);
            ReferencePanelSettings = new ReferencePanelOptions(this);
            wpfControl1.ParentControl = this;
            buttonOpenHide_Click(this, new EventArgs());
            if (!DesignMode)
            {
               // PersistencePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodeEditorControl");
                CompileRequestOptions = new CompileRequestOptions();
                Caret_PositionChanged(this, new EventArgs());
                wpfControl1.TextEditor1.KeyUp += new System.Windows.Input.KeyEventHandler(TextEditor1_KeyUp);
                wpfControl1.TextEditor1.TextArea.Caret.PositionChanged += new EventHandler(Caret_PositionChanged);
            }
            // TODO: focus change if need
        }
         
        #endregion

        #region Events

        /// <summary>
        /// Occurs when a specific key is pressed (see CompileRequestOptions)
        /// </summary>
        [Category("CodeEditor"), Description("Occurs when a specific key is pressed (see CompileRequestOptions)")]
        public event CompileRequestHandler CompileRequest;

        private void RaiseCompileRequest(Key key)
        {
            if (null != CompileRequest)
                CompileRequest(this, new CompileRequestEventArgs(key));
        }

        /// <summary>
        /// Occurs when the text(code) is changed
        /// </summary>
        public new event TextChangedEventHander TextChanged;

        internal void RaiseTextChanged()
        {
            if (null != TextChanged)
                TextChanged(this, new TextChangedEventArgs(wpfControl1.Text));
        }

        #endregion

        #region Properties

        /// <summary>
        /// ErrorPanel settings
        /// </summary>
        [DisplayName("ErrorPanel"), Category("CodeEditor"), Description("ErrorPanel Settings")]
        public ErrorPanelOptions ErrorPanelSettings { get; private set; }

        /// <summary>
        /// ReferencePanel settings
        /// </summary>
        [DisplayName("ReferencePanel"), Category("CodeEditor"), Description("ReferencePanel Settings")]
        public ReferencePanelOptions ReferencePanelSettings { get; private set; }

        /// <summary>
        /// C# Code
        /// </summary>
        [DisplayName("Code"), Category("CodeEditor"), Description("C# Code")]
        public override string Text
        {
            get
            {
                return wpfControl1.Text;
            }
            set
            {
                wpfControl1.Text = value;
            }
        }

        /// <summary>
        /// Specifies whether line numbers are shown on the left to the text view
        /// </summary>
        [DisplayName("ShowLineNumbers"), Category("CodeEditor"), Description("Specifies whether line numbers are shown on the left to the text view")]
        public bool ShowLineNumbers
        {
            get
            {
                return wpfControl1.TextEditor1.ShowLineNumbers;
            }
            set
            {
                wpfControl1.TextEditor1.ShowLineNumbers = value;
            }
        }

        /// <summary>
        /// Assembly info chache path (current codebase if empty)
        /// </summary>
        [Category("CodeEditor"), Description("Assembly info chache path (current codebase if empty)")]
        public string PersistencePath
        {
            get
            {
                return wpfControl1.CurrentFile.PersistancePath;
            }
            set
            {
                wpfControl1.CurrentFile.PersistancePath = value;
            }
        }

    
        /// <summary>
        /// Allows to set a key to fire the CompileRequest event
        /// </summary>
        [DisplayName("RequestOptions"), Category("CodeEditor"), Description("Allows to set a key to fire the CompileRequest event")]
        public CompileRequestOptions CompileRequestOptions { get; set; }

        /// <summary>
        /// info the control is in design mode
        /// </summary>
        [Browsable(false)]
        public new bool DesignMode
        {
            get
            {
                return (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "devenv");
            }
        }
         
        #endregion

        #region Methods
       
        /// <summary>
        /// Add assembly reference from persistence path
        /// </summary>
        /// <param name="assemblyName">Name of the assembly</param>
        /// <param name="doAsync">run as async operation</param>
        public void AddReferenceFromPersistenceFolder(string assemblyName, bool doAsync = false)
        {
            wpfControl1.CurrentFile.AddReferenceFromPersistenceFolder(assemblyName, doAsync);
        }

        /// <summary>
        /// Add assembly references from persistence path in async operation
        /// </summary>
        /// <param name="assemblyName">Name of the assemblies</param>
        /// <param name="doAsync">run as async operation</param>
        public void AddReferencesFromPersistenceFolder(string[] assemblyNames, bool doAsync = false)
        {
            wpfControl1.CurrentFile.AddReferencesFromPersistenceFolder(assemblyNames, doAsync);
        }

        /// <summary>
        /// Add assembly reference from file
        /// </summary>
        /// <param name="assemblyName">Name of the assembly</param>
        /// <param name="assemblyFullPath">Full qualyfied path of the assembly</param>
        /// <param name="tryPersistence">try to find the reference in persistance cache before</param>
        /// <param name="doAsync">run as async operation</param>
        public void AddReferenceFromFile(string assemblyName, string assemblyFullPath, bool tryPersistence = true, bool doAsync = false)
        {
            wpfControl1.CurrentFile.AddReferenceFromFile(assemblyName, assemblyFullPath, tryPersistence, doAsync);
        }

          /// <summary>
        /// Add assembly references from files
        /// </summary>
        /// <param name="assemblyName">Name of the assembly</param>
        /// <param name="assemblyFullPath">Full qualyfied path of the assembly</param>
        /// <param name="tryPersistence">try to find the reference in persistance cache before</param>
        /// <param name="doAsync">run as async operation</param>
        public void AddReferencesFromFile(string[] assemblyNames, string[] assemblyFullPaths, bool tryPersistence = true, bool doAsync = false)
        {
            wpfControl1.CurrentFile.AddReferencesFromFile(assemblyNames, assemblyFullPaths, tryPersistence, doAsync);
        }

        /// <summary>
        /// Set Text property without toogle the TextChanged event
        /// </summary>
        /// <param name="text">new text value</param>
        public void SetTextWithoutChangeEvent(string text)
        {
            wpfControl1.SetTextWithoutChangeEvent(text);
        }

        /// <summary>
        /// Show compiler errors in the panel
        /// </summary>
        /// <param name="errors">error info</param>
        /// <param name="sucseedMessage">optional message if no error occured</param>
        public void ShowErrors(CompilerErrorCollection errors, string sucseedMessage = null)
        {
           if(null == sucseedMessage)
            labelErrors.Text = true == errors.HasErrors ? String.Format("Errors ({0})", errors.Count) : "Errors";
           else
               labelErrors.Text = true == errors.HasErrors ? String.Format("Errors ({0})", errors.Count) : sucseedMessage;
           errorPanel1.ShowErrors(errors);
        }

        /// <summary>
        /// Clear error panel
        /// </summary>
        /// <param name="message">an optional header message</param>
        public void ClearErrors(string message = "Errors")
        {
            labelErrors.Text = message;
            errorPanel1.ClearErrors();
        }

        #endregion

        #region Trigger

        private void Caret_PositionChanged(object sender, EventArgs e)
        {
            int currentLine = wpfControl1.TextEditor1.TextArea.Caret.Line;
            int currentColumn = wpfControl1.TextEditor1.TextArea.Caret.Column;
            labelInfo.Text = String.Format(ErrorPanelSettings.LineInfoFormatString, currentLine, currentColumn);
        }
        
        private void buttonOpenHide_Click(object sender, EventArgs e)
        {
            splitContainer1.Panel2Collapsed = !splitContainer1.Panel2Collapsed;
            buttonOpenHide.Image = true == splitContainer1.Panel2Collapsed ? buttonOpen.Image : buttonHide.Image;
            wpfControl1.TextEditor1.Focus();
        }

        private void referencePanel1_OpenHideClick(object sender, EventArgs e)
        {
            if (referencePanel1.PanelOpen)
            {
                splitContainer3.Panel2MinSize = 100;
                splitContainer3.SplitterWidth = 4;
                splitContainer3.SplitterDistance = this.Width - 100;
                splitContainer3.IsSplitterFixed = false;
                referencePanel1.PerformVisible();                 
            }
            else
            {
                splitContainer3.Panel2MinSize = 10;
                splitContainer3.SplitterWidth = 1;
                splitContainer3.SplitterDistance = this.Width - (17+2);
                splitContainer3.IsSplitterFixed = true;
                referencePanel1.PerformHide();
            }
        }

        private void TextEditor1_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (CompileRequestOptions.Enabled && Convert.ToInt32(e.Key) == Convert.ToInt32(CompileRequestOptions.CompileRequestKey))
                RaiseCompileRequest(CompileRequestOptions.CompileRequestKey);          
        }

        private void errorPanel1_ErrorDoubleClick(ErrorPanel sender, int lineNumber, int columnNumber)
        {
            // another error
            if (0 == lineNumber)
                return;

            string[] split = wpfControl1.TextEditor1.Text.Split(new string[]{Environment.NewLine}, StringSplitOptions.None);
            if (split.Length > lineNumber-1)
            {
                int targetLinePosition = 0;
                for (int i = 0; i < lineNumber -1; i++)
                    targetLinePosition += split[i].Length + Environment.NewLine.Length;
               
                if (split[lineNumber - 1].Length >= columnNumber-1)
                    targetLinePosition += columnNumber - 1;

                wpfControl1.TextEditor1.SelectionStart = targetLinePosition;
                wpfControl1.TextEditor1.Focus();

                int validatedColumnNumber = columnNumber;
                if (split[lineNumber - 1].Length < columnNumber - 1)
                    validatedColumnNumber = 0;

                wpfControl1.TextEditor1.ScrollTo(lineNumber - 1, validatedColumnNumber);
            }
        }

        #endregion
    }
 }
