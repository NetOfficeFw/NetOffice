using System;
using System.Drawing;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ReferencePanelOptions
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">Parent editor control</param>
        internal ReferencePanelOptions(CodeEditorControl parent)
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
        [DisplayName("AllowPanel"), Category("CodeEditor"), Description("Allow the user to see the reference panel")]
        public bool AllowPanel
        {
            get
            {
                return !Parent.splitContainer3.Panel2Collapsed;
            }
            set
            {
                Parent.splitContainer3.Panel2Collapsed = !value;
            }
        }

        /// <summary>
        /// Get or set the error panel is open
        /// </summary>
        [Category("CodeEditor"), Description("Get or set the reference panel is open")]
        public bool PanelOpen
        {
            get
            {
                return Parent.referencePanel1.PanelOpen;
            }
            set
            {
                Parent.referencePanel1.PanelOpen = value;
            }
        }

        /// <summary>
        /// Header message for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Header message for the reference panel")]
        public string Header
        {
            get
            {
                return Parent.referencePanel1.labelHeader.Text;
            }
            set
            {
                Parent.referencePanel1.labelHeader.Text = value;
            }
        }

        /// <summary>
        /// Back color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Back color for the reference panel")]
        public Color BackColor
        {
            get
            {
                return Parent.referencePanel1.BackColor;
            }
            set
            {
                Parent.referencePanel1.BackColor = value;
            }
        }

        /// <summary>
        /// Fore color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Fore color for the reference panel")]
        public Color ForeColor
        {
            get
            {
                return Parent.referencePanel1.ForeColor;
            }
            set
            {
                Parent.referencePanel1.ForeColor = value;
            }
        }

        public override string ToString()
        {
            return "ReferencePanelOptions";
        }
    }
}
