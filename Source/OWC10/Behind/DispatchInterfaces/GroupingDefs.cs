using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface GroupingDefs 
	/// SupportByVersion OWC10, 1
	/// </summary>
	public class GroupingDefs : COMObject, NetOffice.OWC10Api.GroupingDefs
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
                    _contractType = typeof(NetOffice.OWC10Api.GroupingDefs);
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
                    _type = typeof(GroupingDefs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public GroupingDefs() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.OWC10Api.GroupingDef this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.GroupingDef>(this, "Item", typeof(NetOffice.OWC10Api.GroupingDef), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName, object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.GroupingDef>(this, "Add", typeof(NetOffice.OWC10Api.GroupingDef), groupingDefName, groupingFieldName, pageFieldName, index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.GroupingDef>(this, "Add", typeof(NetOffice.OWC10Api.GroupingDef), groupingDefName, groupingFieldName, pageFieldName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType, object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.GroupingDef>(this, "AddTotal", typeof(NetOffice.OWC10Api.GroupingDef), new object[]{ groupingDefName, groupingFieldName, pageFieldName, totalType, index });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.GroupingDef>(this, "AddTotal", typeof(NetOffice.OWC10Api.GroupingDef), groupingDefName, groupingFieldName, pageFieldName, totalType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Delete(object index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.GroupingDef>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.GroupingDef>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.GroupingDef>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.GroupingDef>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual IEnumerator<NetOffice.OWC10Api.GroupingDef> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.GroupingDef item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

