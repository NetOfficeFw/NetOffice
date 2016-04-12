using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Text
{
    public partial class RichTextEditor : UserControl, INotifyPropertyChanged
    {
        #region Fields

        private bool _isInitialized;
        private static PropertyChangedEventArgs _changeArgs = new PropertyChangedEventArgs("RichText");

        #endregion

        #region Ctor

        public RichTextEditor()
        {
            InitializeComponent();
            GetFontCollection();
            _isInitialized = true;
        }
        
        #endregion

        #region Events

        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
        public new event EventHandler TextChanged;

        private void RaiseTextChanged()
        {
            if (null != TextChanged)
                TextChanged(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        public string RichText
        {
            get 
            {
                return richTextBox1.Rtf;
            }
            set
            {
                richTextBox1.Rtf = value;
            }
        }

        #endregion

        #region Methods

        public new bool Focus()
        {
            bool res = base.Focus();
            richTextBox1.Focus();
            return res;
        }

        private void GetFontCollection()
        {
            InstalledFontCollection InsFonts = new InstalledFontCollection();
            foreach (FontFamily item in InsFonts.Families)
                toolStripComboBox1.Items.Add(item.Name);

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void SetSelectionBold()
        {
            Font font = richTextBox1.SelectionFont;
            if (null != font)
            {
                if (font.Bold)
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Regular);
                else
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Bold);
            }
            else             
            {
                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Bold);
            }
        }

        private void SetSelectionItalic()
        {
            Font font = richTextBox1.SelectionFont;
            if (null != font)
            {
                if (font.Italic)
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Regular);
                else
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Italic);
            }
            else
            {

                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Italic);
            }
        }

        private void SetSelectionUnderline()
        {
            Font font = richTextBox1.SelectionFont;
            if (null != font)
            {
                if (font.Underline)
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Regular);
                else
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Underline);
            }
            else
            {
                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Underline);
            }
        }

        private void SetSelectionStrikeout()
        {
            Font font = richTextBox1.SelectionFont;
            if (null != font)
            {
                if (font.Strikeout)
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Regular);
                else
                    richTextBox1.SelectionFont = new Font(font.FontFamily, font.SizeInPoints, FontStyle.Strikeout);
            }
            else
            {
                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Strikeout);
            }
        }

        private void SetSelectionFont()
        {
            if (null != richTextBox1.SelectionFont)
            {
                Font font = richTextBox1.SelectionFont;
                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), font.Size, font.Style);
            }
            else
            {
                richTextBox1.Font = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Regular);
            }
        }

        private void SetSelectionFontSize()
        {
            if(null != richTextBox1.SelectionFont)
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.Name, GetSelectedFontSize(), richTextBox1.SelectionFont.Style);
            else
                richTextBox1.SelectionFont = new Font(toolStripComboBox1.SelectedItem.ToString(), GetSelectedFontSize(), FontStyle.Regular);
        }

        private void SetSelectionFontColor()
        {
            if (colorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                richTextBox1.SelectionColor = colorDialog1.Color;
        }

        private void SetSelectionBackColor()
        {
            if (colorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                richTextBox1.SelectionBackColor = colorDialog1.Color;
        }

        private float GetSelectedFontSize()
        {
            float size;
            if (float.TryParse(toolStripComboBox2.SelectedItem.ToString(), out size))
                return size;
            else
                return 8.0f;
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaiseRichTextChanged()
        {
            if (null != PropertyChanged)
                PropertyChanged(this, _changeArgs);
        }

        #endregion

        #region Trigger

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionFont();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionFontSize();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonBold_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionBold();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonItalic_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionItalic();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonUnderline_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionUnderline();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonStrikeout_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;

                SetSelectionStrikeout();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonForeColor_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionFontColor();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonBackColor_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                SetSelectionBackColor();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                if (null != richTextBox1.SelectionFont)
                {
                    toolStripComboBox2.Text = Convert.ToInt32(richTextBox1.SelectionFont.Size).ToString();
                    toolStripComboBox1.Text = richTextBox1.SelectionFont.Name;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = "Rich Text (*.rtf)|*.rtf|Text (*.txt)|*.txt";
                dlg.InitialDirectory =  Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (DialogResult.OK == dlg.ShowDialog(this))
                    richTextBox1.LoadFile(dlg.FileName);
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (!_isInitialized)
                    return;
                RaiseTextChanged();
                RaiseRichTextChanged();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Rich Text Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
