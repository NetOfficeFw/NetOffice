using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public class AccessContext : BindingList<AccessContextItem>, ITypedList
    {
        internal AccessContext(RootListDefinition parent, string name)
        {
            Name = name;
            Parent = parent;
        }
        
        public string Name{get; private set;}

        private RootListDefinition Parent { get; set; }

        public void CommitLocalChangesToDatabase()
        {

        }

        public void CancelLocalChanges()
        {
        }

        public PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            return Parent.DataSource.GetItemProperties(listAccessors);
        }

        public string GetListName(PropertyDescriptor[] listAccessors)
        {
            return Parent.DataSource.GetListName(listAccessors);
        }
    }
}
