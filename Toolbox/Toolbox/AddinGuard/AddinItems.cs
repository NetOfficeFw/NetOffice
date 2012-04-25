using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using System.ComponentModel;
using System.Text;
using Microsoft.Win32;
//using NetOffice.DeveloperToolbox.RegistryEditor;

namespace NetOffice.DeveloperToolbox.AddinGuard
{
    class AddinItems : IEnumerable
    {
        #region Fields

        WatchController _parent;
        List<AddinsKey> _items = new List<AddinsKey>();

        #endregion
        
        #region Construction

        internal AddinItems(WatchController parent)
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        public AddinsKey Add(string name, RegistryKey rootPath, string registryPath)
        {
            AddinsKey newItem = new AddinsKey(_parent, name, rootPath, registryPath);
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

        public AddinsKey this[int index]
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