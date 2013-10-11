using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Linq;
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
    internal static class Extensions
    {
         public static HighlightingColor GetByName(this IEnumerable<HighlightingColor> collection, string name)
         {
             foreach(HighlightingColor item in collection)
                 if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                     return item;
             return null;
         }
    }
}
