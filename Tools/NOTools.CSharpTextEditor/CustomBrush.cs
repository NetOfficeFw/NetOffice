using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Rendering;

namespace NOTools.CSharpTextEditor
{   
    /// <summary>
    /// Custom brush to change the syntax color
    /// </summary>
    internal sealed class CustomBrush : HighlightingBrush
    {
        private readonly SolidColorBrush _brush;

        public CustomBrush(System.Drawing.Color c)
        {
            var c2 = System.Windows.Media.Color.FromArgb(c.A, c.R, c.G, c.B);
            _brush = CreateFrozenBrush(c2);
        }

        public override System.Windows.Media.Brush GetBrush(ITextRunConstructionContext context)
        {
            return _brush;
        }

        public override string ToString()
        {
            return _brush.ToString();
        }

        private static SolidColorBrush CreateFrozenBrush(System.Windows.Media.Color color)
        {
            SolidColorBrush brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
