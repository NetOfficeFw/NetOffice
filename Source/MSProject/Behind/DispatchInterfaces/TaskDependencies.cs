using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface TaskDependencies 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920720(v=office.14).aspx </remarks>
	public class TaskDependencies : COMObject, NetOffice.MSProjectApi.TaskDependencies
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
                    _contractType = typeof(NetOffice.MSProjectApi.TaskDependencies);
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
                    _type = typeof(TaskDependencies);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TaskDependencies() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.MSProjectApi.TaskDependency this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskDependency>(this, "Item", typeof(NetOffice.MSProjectApi.TaskDependency), index);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Task Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "Parent", typeof(NetOffice.MSProjectApi.Task));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", typeof(NetOffice.MSProjectApi.Application));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Type = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from, object type, object lag)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TaskDependency>(this, "Add", typeof(NetOffice.MSProjectApi.TaskDependency), from, type, lag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TaskDependency>(this, "Add", typeof(NetOffice.MSProjectApi.TaskDependency), from);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Type = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TaskDependency>(this, "Add", typeof(NetOffice.MSProjectApi.TaskDependency), from, type);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.TaskDependency>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.TaskDependency>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.TaskDependency>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.TaskDependency>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        public virtual IEnumerator<NetOffice.MSProjectApi.TaskDependency> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.TaskDependency item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

