using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.SharpDevelop.Dom;

namespace NOTools.CSharpTextEditor
{
    internal class NonFilterStrategy : IFilterStrategy
    {
        /// <summary>
        /// Filters the specified completion items.
        /// </summary>
        /// <param name="completionItems">The completion items.</param>
        public IEnumerable<ICompletionItem> Filter(IEnumerable<ICompletionItem> completionItems)
        {
            return completionItems;
        }
    }
}
