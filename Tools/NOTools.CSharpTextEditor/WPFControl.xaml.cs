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
    /*
                 <codecomplete:CodeCompletionBeahvior 
                                                    FilterStrategy="{Binding FilterStrategy}" 
                                                    ProjectContent="{Binding ProjectContent}" />
     */

    /// <summary>
    /// Interaktionslogik für WPFControl.xaml
    /// </summary>
    public partial class WPFControl : UserControl
    {
        #region Fields

        private TextFile _currentFile;
        private bool _isChange;

        #endregion

        #region Ctor 
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public WPFControl()
        {
            InitializeComponent();
            SetHighlightingAppearance();
            _currentFile = new TextFile();
            DataContext = _currentFile;
            TextEditor1.Document = new ICSharpCode.AvalonEdit.Document.TextDocument();
            TextEditor1.Text = string.Empty;
            SetupContextMenu(); 
        }

        #endregion

        #region Properties
        
        /// <summary>
        /// Control parent
        /// </summary>
        internal CodeEditorControl ParentControl { get; set; }

        /// <summary>
        /// AvalonEdit current TextFile definition 
        /// </summary>
        internal TextFile CurrentFile
        {
            get
            {
                return _currentFile;
            }
        }

        /// <summary>
        /// Gets/Sets the current text
        /// </summary>
        public string Text
        {
            get
            {
                return this.TextEditor1.Text;
            }
            set
            {
                this.TextEditor1.Text = value;
            }
        }

        /// <summary>
        /// Gets/Sets whether to show » for tabs.
        /// </summary>
        public bool ShowTabs
        {
            get 
            {
                return TextEditor1.Options.ShowTabs;
            }
            set
            {
                TextEditor1.Options.ShowTabs = value;
            }
        }

        /// <summary>
        ///  Gets/Sets whether to show · for spaces.
        /// </summary>
        public bool ShowSpaces
        {
            get
            {
                return TextEditor1.Options.ShowSpaces;
            }
            set
            {
                TextEditor1.Options.ShowSpaces = value;
            }
        }

        /// <summary>
        /// Gets/Sets whether to show ¶ at the end of lines.
        /// </summary>
        public bool ShowEndOfLine
        {
            get
            {
                return TextEditor1.Options.ShowEndOfLine;
            }
            set
            {
                TextEditor1.Options.ShowEndOfLine = value;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Set text without toogle the TextChanged event
        /// </summary>
        /// <param name="text"></param>
        public void SetTextWithoutChangeEvent(string text)
        {
            _isChange = true;
            this.TextEditor1.Text = text;
            _isChange = false;
        }

        #endregion

        #region Private Methods

        private void SetupContextMenu()
        {
            CustomContextMenu contextMenu = new CustomContextMenu();
           
            MenuItem itemUndo = new MenuItem();
            itemUndo.Header = "Undo";
            contextMenu.Items.Add(itemUndo);
            itemUndo.Click += new RoutedEventHandler(itemUndo_Click);

            contextMenu.Items.Add(new Separator());

            MenuItem itemCut = new MenuItem();
            itemCut.Header = "Cute";
            itemCut.Click += new RoutedEventHandler(itemCut_Click);
            contextMenu.Items.Add(itemCut);

            MenuItem itemCopy = new MenuItem();
            itemCopy.Header = "Copy";
            itemCopy.Click += new RoutedEventHandler(itemCopy_Click);
            contextMenu.Items.Add(itemCopy);

            MenuItem itemPaste = new MenuItem();
            itemPaste.Header = "Paste";
            itemPaste.Click += new RoutedEventHandler(itemPaste_Click);
            contextMenu.Items.Add(itemPaste);

            MenuItem itemDelete = new MenuItem();
            itemDelete.Header = "Delete";
            itemDelete.Click += new RoutedEventHandler(itemDelete_Click);
            contextMenu.Items.Add(itemDelete);

            contextMenu.Items.Add(new Separator());

            MenuItem itemSelectAll = new MenuItem();
            itemSelectAll.Header = "Select All";
            itemSelectAll.Click += new RoutedEventHandler(itemSelectAll_Click);
            contextMenu.Items.Add(itemSelectAll);

            contextMenu.Opened += new RoutedEventHandler(contextMenu_Opened);
            TextEditor1.ContextMenu = contextMenu;
        }
         
        private void UpdateMenuItemsEnabled()
        {
            CustomContextMenu contextMenu = TextEditor1.ContextMenu as CustomContextMenu;
           
            MenuItem itemUndo = contextMenu.Items[0] as MenuItem;
            itemUndo.IsEnabled = TextEditor1.CanUndo;

            MenuItem itemCut = contextMenu.Items[2] as MenuItem;
            itemCut.IsEnabled = TextEditor1.SelectionLength > 0;
           
            MenuItem itemCopy = contextMenu.Items[3] as MenuItem;
            itemCopy.IsEnabled = TextEditor1.SelectionLength > 0;

            MenuItem itemPaste = contextMenu.Items[4] as MenuItem;
            itemPaste.IsEnabled = !String.IsNullOrEmpty(Clipboard.GetText());

            MenuItem itemDelete = contextMenu.Items[5] as MenuItem;
            itemDelete.IsEnabled = TextEditor1.SelectionLength > 0;

            MenuItem itemSelectAll = contextMenu.Items[7] as MenuItem;
            itemSelectAll.IsEnabled = TextEditor1.Text.Length > 0;
        }

        private void SetHighlightingAppearance()
        {
            TextEditor1.Options.EnableRectangularSelection = false;
            ChangeNamedHighlightingColor("NamespaceKeywords", System.Drawing.Color.Blue);
            ChangeFontWeight("NamespaceKeywords", FontWeights.Normal);
            ChangeNamedHighlightingColor("Keywords", System.Drawing.Color.Blue);
            ChangeFontWeight("Keywords", FontWeights.Normal);
            ChangeNamedHighlightingColor("ThisOrBaseReference", System.Drawing.Color.Blue);
            ChangeFontWeight("ThisOrBaseReference", FontWeights.Normal);
            ChangeNamedHighlightingColor("GetSetAddRemove", System.Drawing.Color.Blue);
            ChangeNamedHighlightingColor("ReferenceTypes", System.Drawing.Color.Blue);
            ChangeNamedHighlightingColor("ValueTypes", System.Drawing.Color.Blue);
            ChangeFontWeight("ValueTypes", FontWeights.Normal);
            ChangeNamedHighlightingColor("Modifiers", System.Drawing.Color.Blue);
            ChangeFontWeight("Modifiers", FontWeights.Normal);
            ChangeNamedHighlightingColor("ParameterModifiers", System.Drawing.Color.Blue);
            ChangeFontWeight("ParameterModifiers", FontWeights.Normal);
            ChangeNamedHighlightingColor("ExceptionKeywords", System.Drawing.Color.Blue);
            ChangeFontWeight("ExceptionKeywords", FontWeights.Normal);
            ChangeNamedHighlightingColor("String", System.Drawing.Color.Maroon);
            ChangeNamedHighlightingColor("Char", System.Drawing.Color.Maroon);
            ChangeNamedHighlightingColor("NullOrValueKeywords", System.Drawing.Color.Blue);
            ChangeFontWeight("NullOrValueKeywords", FontWeights.Normal);
            ChangeNamedHighlightingColor("MethodCall", System.Drawing.Color.Black);
            ChangeFontWeight("MethodCall", FontWeights.Normal);
            ChangeNamedHighlightingColor("TrueFalse", System.Drawing.Color.Blue);
            ChangeFontWeight("TrueFalse", FontWeights.Normal);
            ChangeNamedHighlightingColor("Visibility", System.Drawing.Color.Blue);
            ChangeFontWeight("Visibility", FontWeights.Normal);
            ChangeNamedHighlightingColor("TypeKeywords", System.Drawing.Color.Blue);
            ChangeFontWeight("TypeKeywords", FontWeights.Normal);
            ChangeNamedHighlightingColor("ContextKeywords", System.Drawing.Color.Blue);
            ChangeNamedHighlightingColor("NumberLiteral", System.Drawing.Color.Black);
            ChangeNamedHighlightingColor("Preprocessor", System.Drawing.Color.Blue);
        }

        private void ChangeFontWeight(string colorName, FontWeight weight)
        {   
            HighlightingColor targetColor = TextEditor1.SyntaxHighlighting.NamedHighlightingColors.GetByName(colorName);
            if (null != targetColor)
                targetColor.FontWeight = weight;
            else
                Console.WriteLine("HighlightingColor not found {0}", colorName);
        }

        private void ChangeNamedHighlightingColor(string colorName, System.Drawing.Color color)
        {
            HighlightingColor targetColor = TextEditor1.SyntaxHighlighting.NamedHighlightingColors.GetByName(colorName);
            if (null != targetColor)
                targetColor.Foreground = new CustomBrush(color);
            else
                Console.WriteLine("HighlightingColor not found {0}", colorName);
        }

        #endregion

        #region Trigger

        private void TextEditor1_TextChanged(object sender, EventArgs e)
        {
            if (_isChange)
                return;
            if (null != ParentControl)
                ParentControl.RaiseTextChanged();
        }

        private void contextMenu_Opened(object sender, RoutedEventArgs e)
        {
            UpdateMenuItemsEnabled();
        }

        private void itemSelectAll_Click(object sender, RoutedEventArgs e)
        {
            TextEditor1.SelectAll();
        }

        private void itemDelete_Click(object sender, RoutedEventArgs e)
        {
            if( 0 == TextEditor1.SelectionLength)
                return;
            int startPosition = TextEditor1.SelectionStart;
            int lenght = TextEditor1.SelectionLength;

            string text = TextEditor1.Text;
            string newText = text.Substring(0, startPosition) + text.Substring(startPosition + lenght);
            TextEditor1.Text = newText;

            TextEditor1.SelectionStart = startPosition;
        }

        private void itemPaste_Click(object sender, RoutedEventArgs e)
        {
            string clipBoardText = Clipboard.GetText();
            if (String.IsNullOrEmpty(clipBoardText))                
                return;

            int insertPosition = TextEditor1.SelectionStart;
            string preText = TextEditor1.Text.Substring(0, insertPosition);
            string postText = TextEditor1.Text.Substring(insertPosition);
            TextEditor1.Text = preText + clipBoardText + postText;
        }

        private void itemCut_Click(object sender, RoutedEventArgs e)
        {
            if (0 == TextEditor1.SelectionLength)
                return;
            Clipboard.SetText(TextEditor1.SelectedText);
            itemDelete_Click(sender, e);
        }

        private void itemCopy_Click(object sender, RoutedEventArgs e)
        {
            if (0 == TextEditor1.SelectionLength)
                return;
            Clipboard.SetText(TextEditor1.SelectedText);
        }

        private void itemUndo_Click(object sender, RoutedEventArgs e)
        {
            TextEditor1.Undo();
        }

        private void contextMenu_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            UpdateMenuItemsEnabled();
        }

        #endregion
    }
}
