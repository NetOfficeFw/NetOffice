using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceTasks 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864958.aspx </remarks>
    public class SharedWorkspaceTasks : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SharedWorkspaceTasks
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.SharedWorkspaceTasks);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(SharedWorkspaceTasks);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SharedWorkspaceTasks() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask this[Int32 index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Item", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862401.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862065.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861502.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual bool ItemCountExceeded
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ItemCountExceeded");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        /// <param name="description">optional object description</param>
        /// <param name="dueDate">optional object dueDate</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee, object description, object dueDate)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), new object[] { title, status, priority, assignee, description, dueDate });
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), title);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), title, status);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), title, status, priority);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), title, status, priority, assignee);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        /// <param name="description">optional object description</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee, object description)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceTask>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceTask), new object[] { title, status, priority, assignee, description });
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceTask>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceTask>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceTask>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceTask>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.SharedWorkspaceTask> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.SharedWorkspaceTask item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable 

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
