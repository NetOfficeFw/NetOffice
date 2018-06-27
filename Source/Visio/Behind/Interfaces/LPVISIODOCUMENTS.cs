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
	/// Interface LPVISIODOCUMENTS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	public class LPVISIODOCUMENTS : COMObject, NetOffice.VisioApi.LPVISIODOCUMENTS
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIODOCUMENTS);
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
                    _type = typeof(LPVISIODOCUMENTS);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIODOCUMENTS() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.VisioApi.IVDocument this[object nameOrIndex]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Item", nameOrIndex);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 PersistsEvents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="objectID">Int32 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVDocument get_ItemFromID(Int32 objectID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "ItemFromID", typeof(NetOffice.VisioApi.IVDocument), objectID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="objectID">Int32 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID")]
		public virtual NetOffice.VisioApi.IVDocument ItemFromID(Int32 objectID)
		{
			return get_ItemFromID(objectID);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Add(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "Add", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Open(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "Open", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument OpenEx(string fileName, Int16 flags)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "OpenEx", fileName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nameArray">String[] nameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void GetNames(out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)nameArray);
			Invoker.Method(this, "GetNames", paramsArray, modifiers);
			nameArray = (String[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool CanCheckOut(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void CheckOut(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="measurementSystem">optional NetOffice.VisioApi.Enums.VisMeasurementSystem MeasurementSystem = 0</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument AddEx(string fileName, object measurementSystem, object flags, object langID)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "AddEx", fileName, measurementSystem, flags, langID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDocument AddEx(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "AddEx", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="measurementSystem">optional NetOffice.VisioApi.Enums.VisMeasurementSystem MeasurementSystem = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDocument AddEx(string fileName, object measurementSystem)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "AddEx", fileName, measurementSystem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="measurementSystem">optional NetOffice.VisioApi.Enums.VisMeasurementSystem MeasurementSystem = 0</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVDocument AddEx(string fileName, object measurementSystem, object flags)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVDocument>(this, "AddEx", fileName, measurementSystem, flags);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVDocument>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVDocument>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVDocument>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVDocument>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.VisioApi.IVDocument> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVDocument item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

