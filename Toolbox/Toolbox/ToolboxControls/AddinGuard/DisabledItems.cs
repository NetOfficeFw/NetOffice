using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    class DisabledItems : IEnumerable
    {
        #region Fields

        WatchController _parent;
        List<DisabledKey> _items = new List<DisabledKey>();

        #endregion
        
        #region Construction

        internal DisabledItems(WatchController parent)
        {
            _parent = parent;
        }  
        
        #endregion

        #region Methods

        public DisabledKey Add(string name, RegistryKey rootPath, string registryPath)
        {
            DisabledKey newItem = new DisabledKey(_parent, name, rootPath, registryPath);
            _items.Add(newItem);
            _parent.StopFlag = false;
            return newItem;
        }

        public void Remove(int index)
        {
            _parent.StopFlag = true;
            while (!_parent.StopFlagAgreed)
                ;
            _items.Remove(_items[index]);
            _parent.StopFlag = false;
        }

        public DisabledKey this[int index]
        {
            get
            {
                return _items[index];
            }
        }

        public IEnumerator GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion
    }
}
