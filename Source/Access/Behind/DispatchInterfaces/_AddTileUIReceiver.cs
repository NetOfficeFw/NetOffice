using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _AddTileUIReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _AddTileUIReceiver : COMObject, NetOffice.AccessApi._AddTileUIReceiver
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
                    _contractType = typeof(NetOffice.AccessApi._AddTileUIReceiver);
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
                    _type = typeof(_AddTileUIReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _AddTileUIReceiver() : base()
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
		public virtual string GetClientProtocolVersion()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetClientProtocolVersion");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string CreateCustomTable(string bstrTableName, string bstrNounID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCustomTable", bstrTableName, bstrNounID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetNounsVersion()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetNounsVersion");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetNounsMetadata()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetNounsMetadata");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetDefinitionOfNounID(string bstrNounID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetDefinitionOfNounID", bstrNounID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="pdispNounDefArray">object pdispNounDefArray</param>
		/// <param name="pdispFinalNameArray">object pdispFinalNameArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void CreateObjectFromDefinition(object pdispNounDefArray, object pdispFinalNameArray)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateObjectFromDefinition", pdispNounDefArray, pdispFinalNameArray);
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
		public virtual void CreateRelationship(string bstrLeftTable, string bstrRightTable, string bstrLookupFieldName, string bstrLookupFieldDescription, Int32 lookupFieldPosition, Int32 iOptions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateRelationship", new object[]{ bstrLeftTable, bstrRightTable, bstrLookupFieldName, bstrLookupFieldDescription, lookupFieldPosition, iOptions });
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void ImportData(Int16 type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ImportData", type);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetNounTables()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetNounTables");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrSearchTerm">string bstrSearchTerm</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void RegisterSearchTerm(string bstrSearchTerm)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RegisterSearchTerm", bstrSearchTerm);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void BeginBatchNounAdd()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginBatchNounAdd");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void FinishBatchNounAdd()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FinishBatchNounAdd");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="fVisible">bool fVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void NotifyAddTileUIVisibilityChange(bool fVisible)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NotifyAddTileUIVisibilityChange", fVisible);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void LaunchHyperlink(Int16 type, string bstrUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LaunchHyperlink", type, bstrUrl);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void MetadataLoaded()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MetadataLoaded");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual bool IsOnlineContentAllowed()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsOnlineContentAllowed");
		}

		#endregion

		#pragma warning restore
	}
}

