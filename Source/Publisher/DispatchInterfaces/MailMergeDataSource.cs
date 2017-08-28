using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface MailMergeDataSource 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailMergeDataSource : COMObject
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
                    _type = typeof(MailMergeDataSource);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MailMergeDataSource(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMergeDataSource(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeDataSource(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ActiveRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ActiveRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ActiveRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string ConnectString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectString");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeDataFields DataFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataFields>(this, "DataFields", NetOffice.PublisherApi.MailMergeDataFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeFilters Filters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeFilters>(this, "Filters", NetOffice.PublisherApi.MailMergeFilters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 FirstRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FirstRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool Included
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Included");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Included", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool InvalidAddress
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InvalidAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InvalidAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string InvalidComments
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InvalidComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InvalidComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 LastRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LastRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LastRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeMappedDataFields MappedDataFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeMappedDataFields>(this, "MappedDataFields", NetOffice.PublisherApi.MailMergeMappedDataFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 RecordCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Type
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string TableName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TableName");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeDataSources DataSources
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataSources>(this, "DataSources", NetOffice.PublisherApi.MailMergeDataSources.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsMaster
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool EverValidated
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EverValidated");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EverValidated", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ValidatedClean
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ValidatedClean");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ValidatedClean", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="findText">string findText</param>
		/// <param name="field">optional string Field = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool FindRecord(string findText, object field)
		{
			return Factory.ExecuteBoolMethodGet(this, "FindRecord", findText, field);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="findText">string findText</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public bool FindRecord(string findText)
		{
			return Factory.ExecuteBoolMethodGet(this, "FindRecord", findText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="included">bool included</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetAllIncludedFlags(bool included)
		{
			 Factory.ExecuteMethod(this, "SetAllIncludedFlags", included);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		/// <param name="invalidComment">optional string InvalidComment = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetAllErrorFlags(bool invalid, object invalidComment)
		{
			 Factory.ExecuteMethod(this, "SetAllErrorFlags", invalid, invalidComment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetAllErrorFlags(bool invalid)
		{
			 Factory.ExecuteMethod(this, "SetAllErrorFlags", invalid);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		/// <param name="sortAscending3">optional bool SortAscending3 = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3, sortAscending3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder()
		{
			 Factory.ExecuteMethod(this, "SetSortOrder");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1, object sortAscending1)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1, object sortAscending1, object sortField2)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2, sortAscending2);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void OpenRecipientsDialog()
		{
			 Factory.ExecuteMethod(this, "OpenRecipientsDialog");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyFilter()
		{
			 Factory.ExecuteMethod(this, "ApplyFilter");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lRec">Int32 lRec</param>
		/// <param name="varField">object varField</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void EditRecord(Int32 lRec, object varField, object value)
		{
			 Factory.ExecuteMethod(this, "EditRecord", lRec, varField, value);
		}

		#endregion

		#pragma warning restore
	}
}
