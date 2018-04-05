using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel;
using NetOffice;
using NetOffice.Running;
using NetOffice.CollectionsGeneric;
using NetOffice.Contribution.CollectionsGeneric;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    internal class WindowEntryCollection : SortableBindingList<Entry>, ITypedList, IRefresh
    {
        #region Fields

        private object _lock = new object();

        #endregion

        #region Properties

        public IDisposableSequence<ProxyInformation> WindowItems { get; private set; }

        #endregion

        #region IRefresh

        public bool IsCurrentlyRefresh { get; private set; }

        private Action<IRefresh> Complete { get; set; }

        private Control SyncRoot { get; set; }

        public void RefreshAsync(Action<IRefresh> complete, Control syncRoot)
        {
            if (null == complete)
                throw new ArgumentNullException("complete");
            if (IsCurrentlyRefresh)
                return;

            Complete = complete;
            SyncRoot = syncRoot;
            Action method = Refresh;
            method.BeginInvoke(RefreshCompleted, method);
        }

        private void RefreshCompletedUIThread()
        {
            AddItems();
            RemoveItems();
        }

        private void RefreshCompleted(IAsyncResult result)
        {
            Action method = result.AsyncState as Action;
            try
            {
                method.EndInvoke(result);
                SyncRoot.Invoke(new Action(RefreshCompletedUIThread));
                Complete(this);
            }
            catch
            {
                ;
            }
        }

        private void Refresh()
        {
            lock (_lock)
            {
                try
                {
                    IsCurrentlyRefresh = true;
                    if (null != WindowItems)
                        WindowItems.Dispose();
                    if (Settings.ShowAllAccessible)
                        WindowItems = RunningWindowTable.GetAccessibleProxyInformations(RunningWindowTable.ProxyType.All);
                    else
                        WindowItems = RunningWindowTable.GetAccessibleProxyInformations(RunningWindowTable.ProxyType.AllSupportedOfficeApplications);
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

        public new PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            return TypeDescriptor.GetProperties(typeof(Entry));
        }

        public new string GetListName(PropertyDescriptor[] listAccessors)
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