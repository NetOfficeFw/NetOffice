using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// Interface LPVISIODATARECORDSETS 
	/// SupportByVersion Visio, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIODATARECORDSETS : COMObject ,IEnumerable<NetOffice.VisioApi.IVDataRecordset>
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(LPVISIODATARECORDSETS);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIODATARECORDSETS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATARECORDSETS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.VisioApi.IVDataRecordset this[Int32 index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="iD">Int32 ID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVDataRecordset get_ItemFromID(Int32 iD)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(iD);
			object returnItem = Invoker.PropertyGet(this, "ItemFromID", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="iD">Int32 ID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset ItemFromID(Int32 iD)
		{
			return get_ItemFromID(iD);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EventList", paramsArray);
				NetOffice.VisioApi.IVEventList newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVEventList;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="connectionIDOrString">object ConnectionIDOrString</param>
		/// <param name="commandString">string CommandString</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionIDOrString, commandString, addOptions, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="connectionIDOrString">object ConnectionIDOrString</param>
		/// <param name="commandString">string CommandString</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionIDOrString, commandString, addOptions);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xMLString">string XMLString</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xMLString, addOptions, name);
			object returnItem = Invoker.MethodReturn(this, "AddFromXML", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xMLString">string XMLString</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xMLString, addOptions);
			object returnItem = Invoker.MethodReturn(this, "AddFromXML", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addOptions, name);
			object returnItem = Invoker.MethodReturn(this, "AddFromConnectionFile", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="addOptions">Int32 AddOptions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addOptions);
			object returnItem = Invoker.MethodReturn(this, "AddFromConnectionFile", paramsArray);
			NetOffice.VisioApi.IVDataRecordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataErrorCode">Int32 DataErrorCode</param>
		/// <param name="dataErrorDescription">string DataErrorDescription</param>
		/// <param name="recordsetID">Int32 RecordsetID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetLastDataError(out Int32 dataErrorCode, out string dataErrorDescription, out Int32 recordsetID)
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

       #region IEnumerable<NetOffice.VisioApi.IVDataRecordset> Member
        
        /// <summary>
		/// SupportByVersionAttribute Visio, 12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVDataRecordset> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVDataRecordset item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Visio, 12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}