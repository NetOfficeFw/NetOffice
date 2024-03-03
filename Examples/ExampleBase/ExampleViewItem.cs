using System;
using System.Drawing;

namespace ExampleBase
{
    internal class ExampleViewItem
    {
        internal ExampleViewItem(IExample example, Image icon)
        {
            Caption = example.Caption;
            Description = example.Description;
            Icon = icon;
            Item = example;
        }

        public Image Icon { get; private set; }

        public string Caption { get; private set; }

        public string Description { get; private set; }

        internal IExample Item { get; private set; }
    }
}
