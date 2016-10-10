using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class CommandSyntax
    {
        internal CommandSyntax(Command underlying)
        {
            if (null == underlying)
                throw new ArgumentNullException("underlying");
            Underlying = underlying;
            Items = new CommandsSyntax();
        }

        internal Command Underlying { get; private set; }

        internal CommandsSyntax Items { get; private set; }
    }
}
