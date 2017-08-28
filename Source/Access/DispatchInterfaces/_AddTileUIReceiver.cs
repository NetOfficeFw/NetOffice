using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _AddTileUIReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _AddTileUIReceiver : COMObject
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
                    _type = typeof(_AddTileUIReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _AddTileUIReceiver(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetClientProtocolVersion()
		{
			return Factory.ExecuteStringMethodGet(this, "GetClientProtocolVersion");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string CreateCustomTable(string bstrTableName, string bstrNounID)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCustomTable", bstrTableName, bstrNounID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetNounsVersion()
		{
			return Factory.ExecuteStringMethodGet(this, "GetNounsVersion");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetNounsMetadata()
		{
			return Factory.ExecuteStringMethodGet(this, "GetNounsMetadata");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetDefinitionOfNounID(string bstrNounID)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDefinitionOfNounID", bstrNounID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="pdispNounDefArray">object pdispNounDefArray</param>
		/// <param name="pdispFinalNameArray">object pdispFinalNameArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void CreateObjectFromDefinition(object pdispNounDefArray, object pdispFinalNameArray)
		{
			 Factory.ExecuteMethod(this, "CreateObjectFromDefinition", pdispNounDefArray, pdispFinalNameArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrLeftTable">string bstrLeftTable</param>
		/// <param name="bstrRightTable">string bstrRightTable</param>
		/// <param name="bstrLookupFieldName">string bstrLookupFieldName</param>
		/// <param name="bstrLookupFieldDescription">string bstrLookupFieldDescription</param>
		/// <param name="lookupFieldPosition">Int32 lookupFieldPosition</param>
		/// <param name="iOptions">Int32 iOptions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void CreateRelationship(string bstrLeftTable, string bstrRightTable, string bstrLookupFieldName, string bstrLookupFieldDescription, Int32 lookupFieldPosition, Int32 iOptions)
		{
			 Factory.ExecuteMethod(this, "CreateRelationship", new object[]{ bstrLeftTable, bstrRightTable, bstrLookupFieldName, bstrLookupFieldDescription, lookupFieldPosition, iOptions });
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void ImportData(Int16 type)
		{
			 Factory.ExecuteMethod(this, "ImportData", type);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetNounTables()
		{
			return Factory.ExecuteStringMethodGet(this, "GetNounTables");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrSearchTerm">string bstrSearchTerm</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void RegisterSearchTerm(string bstrSearchTerm)
		{
			 Factory.ExecuteMethod(this, "RegisterSearchTerm", bstrSearchTerm);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void BeginBatchNounAdd()
		{
			 Factory.ExecuteMethod(this, "BeginBatchNounAdd");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void FinishBatchNounAdd()
		{
			 Factory.ExecuteMethod(this, "FinishBatchNounAdd");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="fVisible">bool fVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void NotifyAddTileUIVisibilityChange(bool fVisible)
		{
			 Factory.ExecuteMethod(this, "NotifyAddTileUIVisibilityChange", fVisible);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void LaunchHyperlink(Int16 type, string bstrUrl)
		{
			 Factory.ExecuteMethod(this, "LaunchHyperlink", type, bstrUrl);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void MetadataLoaded()
		{
			 Factory.ExecuteMethod(this, "MetadataLoaded");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public bool IsOnlineContentAllowed()
		{
			return Factory.ExecuteBoolMethodGet(this, "IsOnlineContentAllowed");
		}

		#endregion

		#pragma warning restore
	}
}
