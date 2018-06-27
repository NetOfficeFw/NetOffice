using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIODATARECORDSET 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIODATARECORDSET : COMObject, NetOffice.VisioApi.LPVISIODATARECORDSET
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIODATARECORDSET);
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
                    _type = typeof(LPVISIODATARECORDSET);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIODATARECORDSET() : base()
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
		public virtual Int32 ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisLinkReplaceBehavior LinkReplaceBehavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisLinkReplaceBehavior>(this, "LinkReplaceBehavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LinkReplaceBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDataConnection DataConnection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataConnection>(this, "DataConnection");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDataColumns DataColumns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataColumns>(this, "DataColumns");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual string CommandString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandString");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual string DataAsXML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataAsXML");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual DateTime TimeRefreshed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "TimeRefreshed");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32 RefreshInterval
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RefreshInterval");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RefreshInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32 RefreshSettings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RefreshSettings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RefreshSettings", value);
			}
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
		/// <param name="primaryKeySettings">NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings</param>
		/// <param name="primaryKey">String[] primaryKey</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void GetPrimaryKey(out NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, out String[] primaryKey)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			primaryKeySettings = 0;
			primaryKey = null;
			object[] paramsArray = Invoker.ValidateParamsArray(primaryKeySettings, (object)primaryKey);
			Invoker.Method(this, "GetPrimaryKey", paramsArray, modifiers);
			primaryKeySettings = (NetOffice.VisioApi.Enums.VisPrimaryKeySettings)paramsArray[0];
			primaryKey = (String[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="primaryKeySettings">NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings</param>
		/// <param name="primaryKey">String[] primaryKey</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void SetPrimaryKey(NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, String[] primaryKey)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(primaryKeySettings, (object)primaryKey);
            Invoker.Method(this, "SetPrimaryKey", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="criteriaString">string criteriaString</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32[] GetDataRowIDs(string criteriaString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(criteriaString);
			object returnItem = (object)Invoker.MethodReturn(this, "GetDataRowIDs", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRowID">Int32 dataRowID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual object[] GetRowData(Int32 dataRowID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRowID);
			object returnItem = Invoker.MethodReturn(this, "GetRowData", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
				return newObject;
			}
			else
			{
				return (object[]) returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="newDataAsXML">string newDataAsXML</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void RefreshUsingXML(string newDataAsXML)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshUsingXML", newDataAsXML);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.IVShape[] GetAllRefreshConflicts()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAllRefreshConflicts", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeInConflict">NetOffice.VisioApi.IVShape shapeInConflict</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void RemoveRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveRefreshConflict", shapeInConflict);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeInConflict">NetOffice.VisioApi.IVShape shapeInConflict</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32[] GetMatchingRowsForRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shapeInConflict);
			object returnItem = (object)Invoker.MethodReturn(this, "GetMatchingRowsForRefreshConflict", paramsArray);
			return (Int32[])returnItem;
		}

		#endregion

		#pragma warning restore
	}
}

