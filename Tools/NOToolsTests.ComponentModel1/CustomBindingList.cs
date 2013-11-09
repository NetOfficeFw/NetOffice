using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NOTools.ComponentModel;

namespace NOToolsTests.ComponentModel1
{
    public class CustomBindingList : BindingList<CustomItem>
    {
        protected override void OnBeforeAddInsert(CustomItem item, int itemIndex, ref bool cancel)
        {
            if (item.Name == "New 0 Item")
            {
                Console.WriteLine("CustomBindingList dont want New 0 Item");
                cancel = true;
            }
        }

        protected override void OnAfterAddInsert(CustomItem item, int itemIndex)
        {

        }

        protected override void OnBeforeRemove(CustomItem item, int itemIndex, ref bool cancel)
        {

        }

        protected override void OnAfterRemove(CustomItem item, int itemIndex)
        {

        }
    }
}
