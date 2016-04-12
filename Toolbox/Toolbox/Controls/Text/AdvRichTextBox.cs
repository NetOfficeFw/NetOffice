using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace NetOffice.DeveloperToolbox.Controls.Text
{
    /// <summary>
    /// Represents a standard <see cref="RichTextBox"/> with some
    /// minor added functionality.
    /// </summary>
    /// <remarks>
    /// AdvRichTextBox provides methods to maintain performance
    /// while it is being updated. Additional formatting features
    /// have also been added.
    /// </remarks>
    public class AdvRichTextBox : RichTextBox
    {
        /// <summary>
        /// Maintains performance while updating.
        /// </summary>
        /// <remarks>
        /// <para>
        /// It is recommended to call this method before doing
        /// any major updates that you do not wish the user to
        /// see. Remember to call EndUpdate when you are finished
        /// with the update. Nested calls are supported.
        /// </para>
        /// <para>
        /// Calling this method will prevent redrawing. It will
        /// also setup the event mask of the underlying richedit
        /// control so that no events are sent.
        /// </para>
        /// </remarks>
        public void BeginUpdate()
        {
            // Deal with nested calls.
            ++updating;

            if (updating > 1)
                return;

            // Prevent the control from raising any events.
            oldEventMask = SendMessage(new HandleRef(this, Handle),
                                        EM_SETEVENTMASK, 0, 0);

            // Prevent the control from redrawing itself.
            SendMessage(new HandleRef(this, Handle),
                         WM_SETREDRAW, 0, 0);
        }

        /// <summary>
        /// Resumes drawing and event handling.
        /// </summary>
        /// <remarks>
        /// This method should be called every time a call is made
        /// made to BeginUpdate. It resets the event mask to it's
        /// original value and enables redrawing of the control.
        /// </remarks>
        public void EndUpdate()
        {
            // Deal with nested calls.
            --updating;

            if (updating > 0)
                return;

            // Allow the control to redraw itself.
            SendMessage(new HandleRef(this, Handle),
                         WM_SETREDRAW, 1, 0);

            // Allow the control to raise event messages.
            SendMessage(new HandleRef(this, Handle),
                         EM_SETEVENTMASK, 0, oldEventMask);
        }

        /// <summary>
        /// Gets or sets the alignment to apply to the current
        /// selection or insertion point.
        /// </summary>
        /// <remarks>
        /// Replaces the SelectionAlignment from
        /// <see cref="RichTextBox"/>.
        /// </remarks>
        public new TextAlign SelectionAlignment
        {
            get
            {
                PARAFORMAT fmt = new PARAFORMAT();
                fmt.cbSize = Marshal.SizeOf(fmt);

                // Get the alignment.
                SendMessage(new HandleRef(this, Handle),
                             EM_GETPARAFORMAT,
                             SCF_SELECTION, ref fmt);

                // Default to Left align.
                if ((fmt.dwMask & PFM_ALIGNMENT) == 0)
                    return TextAlign.Left;

                return (TextAlign)fmt.wAlignment;
            }

            set
            {
                PARAFORMAT fmt = new PARAFORMAT();
                fmt.cbSize = Marshal.SizeOf(fmt);
                fmt.dwMask = PFM_ALIGNMENT;
                fmt.wAlignment = (short)value;

                // Set the alignment.
                SendMessage(new HandleRef(this, Handle),
                             EM_SETPARAFORMAT,
                             SCF_SELECTION, ref fmt);
            }
        }

        /// <summary>
        /// This member overrides
        /// <see cref="Control"/>.OnHandleCreated.
        /// </summary>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // Enable support for justification.
            SendMessage(new HandleRef(this, Handle),
                         EM_SETTYPOGRAPHYOPTIONS,
                         TO_ADVANCEDTYPOGRAPHY,
                         TO_ADVANCEDTYPOGRAPHY);
        }

        private int updating = 0;
        private int oldEventMask = 0;

        // Constants from the Platform SDK.
        private const int EM_SETEVENTMASK = 1073;
        private const int EM_GETPARAFORMAT = 1085;
        private const int EM_SETPARAFORMAT = 1095;
        private const int EM_SETTYPOGRAPHYOPTIONS = 1226;
        private const int WM_SETREDRAW = 11;
        private const int TO_ADVANCEDTYPOGRAPHY = 1;
        private const int PFM_ALIGNMENT = 8;
        private const int SCF_SELECTION = 1;

        // It makes no difference if we use PARAFORMAT or
        // PARAFORMAT2 here, so I have opted for PARAFORMAT2.
        [StructLayout(LayoutKind.Sequential)]
        private struct PARAFORMAT
        {
            public int cbSize;
            public uint dwMask;
            public short wNumbering;
            public short wReserved;
            public int dxStartIndent;
            public int dxRightIndent;
            public int dxOffset;
            public short wAlignment;
            public short cTabCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
            public int[] rgxTabs;

            // PARAFORMAT2 from here onwards.
            public int dySpaceBefore;
            public int dySpaceAfter;
            public int dyLineSpacing;
            public short sStyle;
            public byte bLineSpacingRule;
            public byte bOutlineLevel;
            public short wShadingWeight;
            public short wShadingStyle;
            public short wNumberingStart;
            public short wNumberingStyle;
            public short wNumberingTab;
            public short wBorderSpace;
            public short wBorderWidth;
            public short wBorders;
        }

        [DllImport("user32", CharSet = CharSet.Auto)]
        private static extern int SendMessage(HandleRef hWnd,
                                               int msg,
                                               int wParam,
                                               int lParam);

        [DllImport("user32", CharSet = CharSet.Auto)]
        private static extern int SendMessage(HandleRef hWnd,
                                               int msg,
                                               int wParam,
                                               ref PARAFORMAT lp);
    }

    /// <summary>
    /// Specifies how text in a <see cref="AdvRichTextBox"/> is
    /// horizontally aligned.
    /// </summary>
    public enum TextAlign
    {
        /// <summary>
        /// The text is aligned to the left.
        /// </summary>
        Left = 1,

        /// <summary>
        /// The text is aligned to the right.
        /// </summary>
        Right = 2,

        /// <summary>
        /// The text is aligned in the center.
        /// </summary>
        Center = 3,

        /// <summary>
        /// The text is justified.
        /// </summary>
        Justify = 4
    }
}
