using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVDataRecordsets 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	public class IVDataRecordsets : COMObject, NetOffice.VisioApi.IVDataRecordsets
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
                    _contractType = typeof(NetOffice.VisioApi.IVDataRecordsets);
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
                    _type = typeof(IVDataRecordsets);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVDataRecordsets() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.VisioApi.IVDataRecordset this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataRecordset>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVDataRecordset get_ItemFromID(Int32 iD)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVDataRecordset>(this, "ItemFromID", typeof(NetOffice.VisioApi.IVDataRecordset), iD);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("Visio", 12,14,15,16), Redirect("get_ItemFromID")]
		public virtual NetOffice.VisioApi.IVDataRecordset ItemFromID(Int32 iD)
		{
			return get_ItemFromID(iD);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="connectionIDOrString">object connectionIDOrString</param>
		/// <param name="commandString">string commandString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions, object name)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "Add", connectionIDOrString, commandString, addOptions, name);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="connectionIDOrString">object connectionIDOrString</param>
		/// <param name="commandString">string commandString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "Add", connectionIDOrString, commandString, addOptions);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLString">string xMLString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions, object name)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "AddFromXML", xMLString, addOptions, name);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLString">string xMLString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "AddFromXML", xMLString, addOptions);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions, object name)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "AddFromConnectionFile", fileName, addOptions, name);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDataRecordset>(this, "AddFromConnectionFile", fileName, addOptions);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataErrorCode">Int32 dataErrorCode</param>
		/// <param name="dataErrorDescription">string dataErrorDescription</param>
		/// <param name="recordsetID">Int32 recordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void GetLastDataError(out Int32 dataErrorCode, out string dataErrorDescription, out Int32 recordsetID)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true);
			dataErrorCode = 0;
			dataErrorDescription = string.Empty;
			recordsetID = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(dataErrorCode, dataErrorDescription, recordsetID);
			Invoker.Method(this, "GetLastDataError", paramsArray, modifiers);
			dataErrorCode = (Int32)paramsArray[0];
			dataErrorDescription = (string)paramsArray[1];
			recordsetID = (Int32)paramsArray[2];
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVDataRecordset>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVDataRecordset>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVDataRecordset>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVDataRecordset>

        /// <summary>
        /// SupportByVersion Visio, 12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.VisioApi.IVDataRecordset> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVDataRecordset item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

