using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIODATARECORDSET 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIODATARECORDSET : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIODATARECORDSET(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIODATARECORDSET(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSET(string progId) : base(progId)
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
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisLinkReplaceBehavior LinkReplaceBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisLinkReplaceBehavior>(this, "LinkReplaceBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LinkReplaceBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDataConnection DataConnection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataConnection>(this, "DataConnection");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDataColumns DataColumns
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataColumns>(this, "DataColumns");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public string CommandString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public string DataAsXML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataAsXML");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public DateTime TimeRefreshed
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "TimeRefreshed");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 RefreshInterval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RefreshInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 RefreshSettings
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RefreshSettings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshSettings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
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
		public void GetPrimaryKey(out NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, out String[] primaryKey)
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
		public void SetPrimaryKey(NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, String[] primaryKey)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(primaryKeySettings, (object)primaryKey);
            Invoker.Method(this, "SetPrimaryKey", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="criteriaString">string criteriaString</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32[] GetDataRowIDs(string criteriaString)
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
		public object[] GetRowData(Int32 dataRowID)
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
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="newDataAsXML">string newDataAsXML</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void RefreshUsingXML(string newDataAsXML)
		{
			 Factory.ExecuteMethod(this, "RefreshUsingXML", newDataAsXML);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVShape[] GetAllRefreshConflicts()
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
		public void RemoveRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict)
		{
			 Factory.ExecuteMethod(this, "RemoveRefreshConflict", shapeInConflict);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeInConflict">NetOffice.VisioApi.IVShape shapeInConflict</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32[] GetMatchingRowsForRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shapeInConflict);
			object returnItem = (object)Invoker.MethodReturn(this, "GetMatchingRowsForRefreshConflict", paramsArray);
			return (Int32[])returnItem;
		}

		#endregion

		#pragma warning restore
	}
}
