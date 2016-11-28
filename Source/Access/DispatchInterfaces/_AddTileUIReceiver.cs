using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _AddTileUIReceiver 
	/// SupportByVersion Access, 15, 16
	///</summary>
	[SupportByVersionAttribute("Access", 15, 16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _AddTileUIReceiver : COMObject
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
                    _type = typeof(_AddTileUIReceiver);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _AddTileUIReceiver(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AddTileUIReceiver(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetClientProtocolVersion()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetClientProtocolVersion", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string CreateCustomTable(string bstrTableName, string bstrNounID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrTableName, bstrNounID);
			object returnItem = Invoker.MethodReturn(this, "CreateCustomTable", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetNounsVersion()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetNounsVersion", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetNounsMetadata()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetNounsMetadata", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetDefinitionOfNounID(string bstrNounID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNounID);
			object returnItem = Invoker.MethodReturn(this, "GetDefinitionOfNounID", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="pdispNounDefArray">object pdispNounDefArray</param>
		/// <param name="pdispFinalNameArray">object pdispFinalNameArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void CreateObjectFromDefinition(object pdispNounDefArray, object pdispFinalNameArray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pdispNounDefArray, pdispFinalNameArray);
			Invoker.Method(this, "CreateObjectFromDefinition", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrLeftTable">string bstrLeftTable</param>
		/// <param name="bstrRightTable">string bstrRightTable</param>
		/// <param name="bstrLookupFieldName">string bstrLookupFieldName</param>
		/// <param name="bstrLookupFieldDescription">string bstrLookupFieldDescription</param>
		/// <param name="lookupFieldPosition">Int32 lookupFieldPosition</param>
		/// <param name="iOptions">Int32 iOptions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void CreateRelationship(string bstrLeftTable, string bstrRightTable, string bstrLookupFieldName, string bstrLookupFieldDescription, Int32 lookupFieldPosition, Int32 iOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrLeftTable, bstrRightTable, bstrLookupFieldName, bstrLookupFieldDescription, lookupFieldPosition, iOptions);
			Invoker.Method(this, "CreateRelationship", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="type">Int16 Type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void ImportData(Int16 type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ImportData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetNounTables()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetNounTables", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrSearchTerm">string bstrSearchTerm</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void RegisterSearchTerm(string bstrSearchTerm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSearchTerm);
			Invoker.Method(this, "RegisterSearchTerm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void BeginBatchNounAdd()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BeginBatchNounAdd", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void FinishBatchNounAdd()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FinishBatchNounAdd", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="fVisible">bool fVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void NotifyAddTileUIVisibilityChange(bool fVisible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fVisible);
			Invoker.Method(this, "NotifyAddTileUIVisibilityChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="type">Int16 Type</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void LaunchHyperlink(Int16 type, string bstrUrl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, bstrUrl);
			Invoker.Method(this, "LaunchHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void MetadataLoaded()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MetadataLoaded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public bool IsOnlineContentAllowed()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "IsOnlineContentAllowed", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}