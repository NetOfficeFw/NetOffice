using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;

namespace ProxyView
{
    internal class RotEntryCollection : SortableBindingList<Entry>, ITypedList, IRefresh
    {
        #region Fields

        private object _lock = new object();

        #endregion

        #region Properties

        public bool IsCurrentlyRefresh { get; private set; }

        private IDisposableEnumeration<ProxyInformation> RotItems { get; set; }

        #endregion

        #region Methods

        public void Refresh()
        {
            try
            {
                lock (_lock)
                {
                    IsCurrentlyRefresh = true;
                    if (null != RotItems)
                        RotItems.Dispose();
                    RotItems = RunningObjectTable.GetActiveProxyInformations("", "");
                    AddItems();
                    RemoveItems();                   
                }                
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsCurrentlyRefresh = false;
            }           
        }

        private void AddItems()
        {
            foreach (var item in RotItems)
            {
                if (!Contains(item))
                {
                    Add(new Entry(
                        item.Proxy,
                        item.ID,
                        item.DisplayName,
                        item.Name,
                        item.Component,
                        item.Library,
                        item.ProcessID.ToString() + TryGetProcessName(item.ProcessID),
                        item.Elevation
                        ));
                }
            }
        }

        private static string TryGetProcessName(IntPtr processID)
        {
            if (processID == IntPtr.Zero)
                return String.Empty;

            System.Diagnostics.Process proc = System.Diagnostics.Process.GetProcessById(processID.ToInt32());
            if (null != proc)
                return "(" + proc.ProcessName + ")";
            else
                return String.Empty;
        }

        private void RemoveItems()
        {
            List<Entry> itemsToDelete = new List<Entry>();
            foreach (var item in this)
            {
                if (!RotItemsContains(item.Underlying))
                    itemsToDelete.Add(item);
            }

            foreach (var item in itemsToDelete)
                Remove(item);
        }

        private bool RotItemsContains(object comProxy)
        {
            foreach (var item in RotItems)
            {
                if (item.Proxy == comProxy)
                    return true;
            }
            return false;
        }

        private bool Contains(ProxyInformation proxyInfo)
        {
            foreach (var item in this)
            {
                if (item.ID == proxyInfo.ID && item.Caption == proxyInfo.DisplayName &&
                    item.Component == item.Component && item.Library == proxyInfo.Library)
                    return true;
            }
            return false;
        }
          
        #endregion

        #region ITypedList

        public PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            return TypeDescriptor.GetProperties(typeof(Entry));
        }

        public string GetListName(PropertyDescriptor[] listAccessors)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            if (null != RotItems)
                RotItems.Dispose();
        }

        #endregion
    }
}
