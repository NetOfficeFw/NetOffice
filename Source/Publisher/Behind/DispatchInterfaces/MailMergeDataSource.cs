using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface MailMergeDataSource 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailMergeDataSource : COMObject, NetOffice.PublisherApi.MailMergeDataSource
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
                    _contractType = typeof(NetOffice.PublisherApi.MailMergeDataSource);
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
                    _type = typeof(MailMergeDataSource);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailMergeDataSource() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 ActiveRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ActiveRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ActiveRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string ConnectString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectString");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeDataFields DataFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataFields>(this, "DataFields", typeof(NetOffice.PublisherApi.MailMergeDataFields));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeFilters Filters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeFilters>(this, "Filters", typeof(NetOffice.PublisherApi.MailMergeFilters));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 FirstRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FirstRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FirstRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool Included
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Included");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Included", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool InvalidAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InvalidAddress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvalidAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string InvalidComments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "InvalidComments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvalidComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 LastRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LastRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LastRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeMappedDataFields MappedDataFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeMappedDataFields>(this, "MappedDataFields", typeof(NetOffice.PublisherApi.MailMergeMappedDataFields));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 RecordCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string TableName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TableName");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeDataSources DataSources
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataSources>(this, "DataSources", typeof(NetOffice.PublisherApi.MailMergeDataSources));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool IsMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool EverValidated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EverValidated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EverValidated", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ValidatedClean
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ValidatedClean");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ValidatedClean", value);
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
		public virtual bool FindRecord(string findText, object field)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FindRecord", findText, field);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="findText">string findText</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool FindRecord(string findText)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FindRecord", findText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="included">bool included</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetAllIncludedFlags(bool included)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAllIncludedFlags", included);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		/// <param name="invalidComment">optional string InvalidComment = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetAllErrorFlags(bool invalid, object invalidComment)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAllErrorFlags", invalid, invalidComment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetAllErrorFlags(bool invalid)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAllErrorFlags", invalid);
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
		public virtual void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3, sortAscending3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetSortOrder()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetSortOrder(object sortField1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetSortOrder(object sortField1, object sortAscending1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetSortOrder(object sortField1, object sortAscending1, object sortField2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2);
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
		public virtual void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2, sortAscending2);
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
		public virtual void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenRecipientsDialog()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenRecipientsDialog");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyFilter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lRec">Int32 lRec</param>
		/// <param name="varField">object varField</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EditRecord(Int32 lRec, object varField, object value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditRecord", lRec, varField, value);
		}

		#endregion

		#pragma warning restore
	}
}


