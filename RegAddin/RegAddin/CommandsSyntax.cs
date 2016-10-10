using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class CommandsSyntax : IEnumerable<CommandSyntax>
    {
        private List<CommandSyntax> _items = new List<CommandSyntax>();

        internal void Add(CommandSyntax info)
        {
            if (null == info)
                throw new ArgumentNullException("info");
            _items.Add(info);
        }

        public IEnumerator<CommandSyntax> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }
    }
}
