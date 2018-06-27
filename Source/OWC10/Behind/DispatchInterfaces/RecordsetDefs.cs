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
	/// DispatchInterface RecordsetDefs 
	/// SupportByVersion OWC10, 1
	/// </summary>
	public class RecordsetDefs : COMObject, NetOffice.OWC10Api.RecordsetDefs
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
                    _contractType = typeof(NetOffice.OWC10Api.RecordsetDefs);
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
                    _type = typeof(RecordsetDefs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public RecordsetDefs() : base()
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
		public virtual NetOffice.OWC10Api.RecordsetDef this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDef>(this, "Item", typeof(NetOffice.OWC10Api.RecordsetDef), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="schemaRowsource">object schemaRowsource</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "Add", typeof(NetOffice.OWC10Api.RecordsetDef), schemaRowsource, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="schemaRowsource">object schemaRowsource</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "Add", typeof(NetOffice.OWC10Api.RecordsetDef), schemaRowsource);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="rowsourceType">optional object rowsourceType</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "AddNew", typeof(NetOffice.OWC10Api.RecordsetDef), source, rowsourceType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.RecordsetDef AddNew(string source)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "AddNew", typeof(NetOffice.OWC10Api.RecordsetDef), source);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="rowsourceType">optional object rowsourceType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "AddNew", typeof(NetOffice.OWC10Api.RecordsetDef), source, rowsourceType);
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

        #region IEnumerableProvider<NetOffice.OWC10Api.RecordsetDef>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.RecordsetDef>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.RecordsetDef>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.RecordsetDef>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual IEnumerator<NetOffice.OWC10Api.RecordsetDef> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.RecordsetDef item in innerEnumerator)
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

