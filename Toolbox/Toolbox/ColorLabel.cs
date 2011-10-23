using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public enum ColorShift
    { 
        Red=0,
        Green=1,
        Blue =2,
        All = 3        
    }

    public partial class ColorLabel : UserControl
    {
        #region Fields

        Color _backColor = Color.LightGray;
        Color _foreColor = Color.DarkGray;
        Color _effectForeColor = Color.White;

        int _colorForeShiftFactor = 10;
        int _colorBackShiftFactor = -5;
        ColorShift _colorShift = ColorShift.All;
        int _currentEffectPosition;

        #endregion
        
        #region Construction

        public ColorLabel()
        {
            InitializeComponent();
            this.Text = "NetOffice Developer Toolbox";
            // if (!DesignMode)
            //    StartEffectTimer();
        }
        
        #endregion

        #region Properties
       
        private new bool DesignMode
        {
            get
            {
                return (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "devenv");
            }
        }

        public ColorShift ColorShift
        {
            get
            {
                return _colorShift;
            }
            set
            {
                _colorShift = value;
            }
        }

        public int ColorBackShiftFactor
        {
            get             
            {
                return _colorBackShiftFactor;
            }
            set 
            {
                _colorBackShiftFactor = value;
            }
        }

        public new Color BackColor
        {
            get
            {
                return _backColor;
            }
            set
            {
                _backColor = value;
            }
        }

        public new string Text
        {
             
            get 
            {
                return GetText();
            }
            set 
            {
                SetupText(value);
                SetupSize();
            }
        }

        #endregion

        #region Methods
       
        private void SetupSize()
        {
            this.Height = 20;
            this.Width = GetNextLeftPosition();
        }

        private void StartEffectTimer()
        {
            _currentEffectPosition = -1;
            timerEffect.Enabled = true;
        }

        private string GetText()
        {
            string result = "";
            foreach (Control item in this.Controls)
                result += item.Text;
            return result;
        }

        private void SetupText(string text)
        {
            this.Controls.Clear();
            int i = 0;
            foreach (char item in text.ToCharArray())
            {
                Label newLabel = new Label();
               
                newLabel.Top = 0;
                newLabel.Left = GetNextLeftPosition()-6;
                newLabel.Name = string.Format("Label{0}", i.ToString());
                newLabel.Font = new System.Drawing.Font("Microsoft Sans Sherif", 16, FontStyle.Bold);
                newLabel.BackColor = GetCurrentBackColor(i);
                newLabel.ForeColor = GetCurrentForeColor(i);
                newLabel.Text = item.ToString();
                newLabel.TextAlign = ContentAlignment.BottomLeft;
                newLabel.AutoSize = true; 
                this.Controls.Add(newLabel);
                i++;
            }
        }

        private Color GetCurrentBackColor(int index)
        {
            switch (_colorShift)
            {
                case DeveloperToolbox.ColorShift.Red:
                    return ShiftRed(index, _backColor, _colorBackShiftFactor);
                case DeveloperToolbox.ColorShift.Green:
                    return ShiftGreen(index, _backColor, _colorBackShiftFactor);
                case DeveloperToolbox.ColorShift.Blue:
                    return ShiftBlue(index, _backColor, _colorBackShiftFactor);
                case DeveloperToolbox.ColorShift.All:
                    return ShiftAll(index, _backColor, _colorBackShiftFactor);
            }

            throw new InvalidEnumArgumentException("_colorShift is unkown.");
        }

        private Color GetCurrentForeColor(int index)
        {
            int shiftFactor = _colorForeShiftFactor;
            switch (index)
            {
                case 2:
                case 3:
                    shiftFactor -= (shiftFactor / 2);
                    break;
            }
               
            switch (_colorShift)
            {
                case DeveloperToolbox.ColorShift.Red:
                    return ShiftRed(index, _foreColor, shiftFactor);
                case DeveloperToolbox.ColorShift.Green:
                    return ShiftGreen(index, _foreColor, shiftFactor);
                case DeveloperToolbox.ColorShift.Blue:
                    return ShiftBlue(index, _foreColor, shiftFactor);
                case DeveloperToolbox.ColorShift.All:
                    return ShiftAll(index, _foreColor, shiftFactor);
            }

            throw new InvalidEnumArgumentException("_colorShift is unkown.");
        }

        private Color ShiftAll(int index, Color color, int shiftFactor)
        {
            byte blue = color.B;
            int newblue = ((int)blue) + (shiftFactor * index);
            if (newblue > 255) newblue = 255;
            if (newblue < 0) newblue = 0;

            byte green = color.G;
            int newGreen = ((int)green) + (shiftFactor * index);
            if (newGreen > 255) newGreen = 255;
            if (newGreen < 0) newGreen = 0;

            byte red = color.R;
            int newRed = ((int)red) + (shiftFactor * index);
            if (newRed > 255) newRed = 255;
            if (newRed < 0) newRed = 0;

            return Color.FromArgb(newRed, newGreen, newblue);
        }

        private Color ShiftBlue(int index, Color color, int shiftFactor)
        {
            byte blue = color.B;
            int newblue = ((int)blue) + (shiftFactor * index);
            if (newblue > 255) newblue = 255;
            if (newblue < 0) newblue = 0;
            return Color.FromArgb(color.R, color.G, newblue);
        }

        private Color ShiftGreen(int index, Color color, int shiftFactor)
        {
            byte green = color.G;
            int newGreen = ((int)green) + (shiftFactor * index);
            if (newGreen > 255) newGreen = 255;
            if (newGreen < 0) newGreen = 0;
            return Color.FromArgb(color.R, newGreen, color.B);
        }

        private Color ShiftRed(int index, Color color, int shiftFactor)
        {
            byte red = color.R;
            int newRed = ((int)red) + (shiftFactor * index);
            if (newRed > 255) newRed = 255;
            if (newRed < 0) newRed = 0;
            return Color.FromArgb(newRed, color.G, color.B);
        }

        private int GetNextLeftPosition()
        { 
            if(this.Controls.Count > 0)
            {
                Label control = this.Controls[this.Controls.Count - 1] as Label;
                return control.Left + control.Width + 1;             
            }
            else
                return 0;                            
        }
        
        #endregion
        
        #region Trigger
        
        private void timerEffect_Tick(object sender, EventArgs e)
        {
            int maxCount = this.Controls.Count;
            if (_currentEffectPosition < 0)
            {
                _currentEffectPosition = 0;
                Label newEffectControl = this.Controls[_currentEffectPosition] as Label;
                newEffectControl.Tag = newEffectControl.ForeColor;
                newEffectControl.ForeColor = _effectForeColor;
                _currentEffectPosition++;
            }
            else if (_currentEffectPosition == maxCount)
            {
                Label oldEffectControl = this.Controls[_currentEffectPosition - 1] as Label;
                oldEffectControl.ForeColor = (Color)oldEffectControl.Tag;
                Label newEffectControl = this.Controls[0] as Label;
                newEffectControl.Tag = newEffectControl.ForeColor;
                newEffectControl.ForeColor = _effectForeColor;
                _currentEffectPosition=1;
            }
            else
            {
                Label oldEffectControl = this.Controls[_currentEffectPosition-1] as Label;
                oldEffectControl.ForeColor = (Color)oldEffectControl.Tag;
                Label newEffectControl = this.Controls[_currentEffectPosition] as Label;
                newEffectControl.Tag = newEffectControl.ForeColor;
                newEffectControl.ForeColor = _effectForeColor;
                _currentEffectPosition++;
            }
        }
        
        #endregion
    }
}
