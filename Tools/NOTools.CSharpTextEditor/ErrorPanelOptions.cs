using System;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    /// <summary>
    /// All possible settings for the ErrorPanel
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ErrorPanelOptions
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">Parent editor control</param>
        internal ErrorPanelOptions(CodeEditorControl parent)
        {
            Parent = parent;
        }

        /// <summary>
        /// Parent editor control
        /// </summary>
        private CodeEditorControl Parent { get; set; }

        /// <summary>
        /// Allow the user to see the error panel
        /// </summary>
        [DisplayName("AllowPanel"), Category("CodeEditor"), Description("Allow the user to see the error panel")]
        public bool AllowPanel
        {
            get
            {
                return Parent.buttonErrorPanelOpenHide.Visible;
            }
            set
            {

                Parent.buttonErrorPanelOpenHide.Visible = value;
            }
        }

        /// <summary>
        /// Get or set the error panel is open
        /// </summary>
        [Category("CodeEditor"), Description("Get or set the error panel is open")]
        public bool PanelOpen
        {
            get
            {
                return !Parent.splitContainer1.Panel2Collapsed;
            }
            set
            {
                Parent.splitContainer1.Panel2Collapsed = !value;
                Parent.buttonErrorPanelOpenHide.Image = true == Parent.splitContainer1.Panel2Collapsed ? Parent.buttonOpen.Image : Parent.buttonHide.Image;
                Parent.wpfControl1.TextEditor1.Focus();
            }
        }

        /// <summary>
        /// Header message for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Header message for the error panel")]
        public string Header
        {
            get
            {
                return Parent.labelErrors.Text;
            }
            set
            {
                Parent.labelErrors.Text = value;
            }
        }

        /// <summary>
        /// Line info format string for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Line info format string for the error panel")]
        public string LineInfoFormatString
        {
            get
            {
                return _lineInfoFormatString;
            }
            set
            {
                _lineInfoFormatString = value;
            }
        }
        private string _lineInfoFormatString = "Current Line:{0} Position:{1}";

        /// <summary>
        /// Column header for the first column
        /// </summary>
        [Category("CodeEditor"), Description("Column header for the first column")]
        public string LineColumnHeader
        {
            get
            {
                return Parent.errorPanel1.LineColumnHeader;
            }
            set
            {
                Parent.errorPanel1.LineColumnHeader = value;
            }
        }

        /// <summary>
        /// Column header for the second column
        /// </summary>
        [Category("CodeEditor"), Description("Column header for the second column")]
        public string ErrorColumnHeader
        {
            get
            {
                return Parent.errorPanel1.ErrorColumnHeader;
            }
            set
            {
                Parent.errorPanel1.ErrorColumnHeader = value;
            }
        }

        /// <summary>
        /// Back color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Back color for the error panel")]
        public Color BackColor
        {
            get
            {
                return _panelBackColor;
            }
            set
            {
                Parent.panelHeader.BackColor = value;
                Parent.errorPanel1.SetColumnColor(value);
                _panelBackColor = value;
            }
        }
        private Color _panelBackColor = Color.LightSteelBlue;

        /// <summary>
        /// Fore color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Fore color for the error panel")]
        public Color ForeColor
        {
            get
            {
                return Parent.labelErrors.ForeColor;
            }
            set
            {
                Parent.labelErrors.ForeColor = value;
                Parent.labelInfo.ForeColor = value;
            }
        }

        public override string ToString()
        {
            return "ErrorPanelOptions";
        }
    }

}
