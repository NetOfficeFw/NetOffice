using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.Running;
using NetOffice.CollectionsGeneric;
using NetOffice.Contribution.CollectionsGeneric;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    internal class RotEntryCollection : SortableBindingList<Entry>, ITypedList, IRefresh
    {
        #region Fields

        private object _lock = new object();

        #endregion

        #region Properties

        private IDisposableSequence<ProxyInformation> RotItems { get; set; }

        #endregion

        #region Methods

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

        private void RemoveItems()
        {
            List<Entry> itemsToDelete = new List<Entry>();
            foreach (var item in this)
            {
                if (!RotItemsContains(item))
                    itemsToDelete.Add(item);
            }

            foreach (var item in itemsToDelete)
                Remove(item);
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

        private bool RotItemsContains(Entry proxyInfo)
        {
            foreach (var item in RotItems)
            {
                if (item.ID == proxyInfo.ID && item.DisplayName == proxyInfo.Caption &&
                  item.Component == item.Component && (String.IsNullOrWhiteSpace(item.Library) ? "<Unknown>" : item.Library) == proxyInfo.Library)
                    return true;
            }
            return false;
        }

        private bool Contains(ProxyInformation proxyInfo)
        {
            foreach (var item in this)
            {
                if (item.ID == proxyInfo.ID && item.Caption == proxyInfo.DisplayName &&
                    item.Component == item.Component && item.Library == (String.IsNullOrWhiteSpace(proxyInfo.Library) ? "<Unknown>" : proxyInfo.Library))
                    return true;
            }
            return false;
        }

        #endregion

        #region IRefresh

        public bool IsCurrentlyRefresh { get; private set; }

        private Action<IRefresh> Complete { get; set; }

        private Control SyncRoot { get; set; }

        public void RefreshAsync(Action<IRefresh> complete, Control syncRoot)
        {
            if (null == complete)
                throw new ArgumentNullException("complete");
            if (null == syncRoot)
                throw new ArgumentNullException("syncRoot");
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
            try
            {
                lock (_lock)
                {
                    IsCurrentlyRefresh = true;
                    if (null != RotItems)
                        RotItems.Dispose();
                    RotItems = RunningObjectTable.GetActiveProxyInformations("", "");

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

        public void Dispose()
        {
            if (null != RotItems)
                RotItems.Dispose();
        }

        #endregion
    }
}
