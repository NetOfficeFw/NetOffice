using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel;
using NetOffice;

namespace ProxyView
{
    internal class WindowEntryCollection : SortableBindingList<Entry>, ITypedList, IRefresh
    {
        #region Fields

        private object _lock = new object();

        #endregion

        #region Properties

        public IDisposableEnumeration<ProxyInformation> WindowItems { get; private set; }

        #endregion

        #region IRefresh

        public bool IsCurrentlyRefresh { get; private set; }

        public void Refresh()
        {
            lock (_lock)
            {
                try
                {
                    IsCurrentlyRefresh = true;
                    if (null != WindowItems)
                        WindowItems.Dispose();
                    if (Properties.Settings.Default.ShowAllAccessible)
                        WindowItems = RunningWindowTable.GetAccessibleProxyInformations(RunningWindowTable.ProxyType.All);
                    else
                        WindowItems = RunningWindowTable.GetAccessibleProxyInformations(RunningWindowTable.ProxyType.AllSupportedOfficeApplications);
                    AddItems();
                    RemoveItems();
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
           
        }

        #endregion

        #region Methods
          
        private void AddItems()
        {
            foreach (var item in WindowItems)
            {
                if (!Contains(item.Proxy))
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
                if (!WndItemsContains(item.Underlying))
                    itemsToDelete.Add(item);
            }

            foreach (var item in itemsToDelete)
                Remove(item);
        }

        private bool WndItemsContains(object comProxy)
        {
            foreach (var item in WindowItems)
            {
                if (item.Proxy == comProxy)
                    return true;
            }
            return false;
        }

        private bool Contains(object comProxy)
        {
            foreach (var item in this)
            {
                if (item.Underlying == comProxy)
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
        
        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (null != WindowItems)
                WindowItems.Dispose();
        }

        #endregion
    }
}
