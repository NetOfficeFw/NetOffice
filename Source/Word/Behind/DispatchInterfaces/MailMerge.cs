using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836701.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailMerge : COMObject, NetOffice.WordApi.MailMerge
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
                    _contractType = typeof(NetOffice.WordApi.MailMerge);
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
                    _type = typeof(MailMerge);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailMerge() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837890.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837179.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839335.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdMailMergeMainDocType MainDocumentType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeMainDocType>(this, "MainDocumentType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MainDocumentType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840195.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdMailMergeState State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845069.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdMailMergeDestination Destination
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeDestination>(this, "Destination");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Destination", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838518.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeDataSource DataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMergeDataSource>(this, "DataSource", typeof(NetOffice.WordApi.MailMergeDataSource));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193129.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeFields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMergeFields>(this, "Fields", typeof(NetOffice.WordApi.MailMergeFields));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840472.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 ViewMailMergeFieldCodes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ViewMailMergeFieldCodes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewMailMergeFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192581.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SuppressBlankLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SuppressBlankLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SuppressBlankLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841091.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MailAsAttachment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MailAsAttachment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MailAsAttachment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820768.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string MailAddressFieldName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MailAddressFieldName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MailAddressFieldName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820986.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string MailSubject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MailSubject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MailSubject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845591.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool HighlightMergeFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HighlightMergeFields");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HighlightMergeFields", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192784.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdMailMergeMailFormat MailFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeMailFormat>(this, "MailFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MailFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192539.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual string ShowSendToCustom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShowSendToCustom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSendToCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Int32 WizardState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WizardState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WizardState", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		/// <param name="mSQuery">optional object mSQuery</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection, object linkToSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection, linkToSource });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", name, passwordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", name, passwordDocument, writePasswordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", name, passwordDocument, writePasswordDocument, headerRecord);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		/// <param name="mSQuery">optional object mSQuery</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		/// <param name="mSQuery">optional object mSQuery</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		/// <param name="mSQuery">optional object mSQuery</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		/// <param name="mSQuery">optional object mSQuery</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="connection">optional object connection</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument, writePasswordDocument, headerRecord);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateHeaderSource(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateHeaderSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateHeaderSource(string name, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument, writePasswordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="openExclusive">optional object openExclusive</param>
		/// <param name="subType">optional object subType</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive, object subType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive, subType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", name, format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", name, format, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", name, format, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="openExclusive">optional object openExclusive</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="openExclusive">optional object openExclusive</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object openExclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, openExclusive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", name, format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", name, format, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", name, format, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		/// <param name="pause">optional object pause</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Execute(object pause)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", pause);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Execute()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835814.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Check()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Check");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192805.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditDataSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditDataSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838561.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditHeaderSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditHeaderSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845149.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditMainDocument()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditMainDocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">string type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UseAddressBook(string type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UseAddressBook", type);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", name, format);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", name, format, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", name, format, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", name, format);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", name, format, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", name, format, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		/// <param name="showDataStep">optional object showDataStep</param>
		/// <param name="showWriteStep">optional object showWriteStep</param>
		/// <param name="showPreviewStep">optional object showPreviewStep</param>
		/// <param name="showMergeStep">optional object showMergeStep</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", initialState);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep, showTemplateStep);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		/// <param name="showDataStep">optional object showDataStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep, showTemplateStep, showDataStep);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		/// <param name="showDataStep">optional object showDataStep</param>
		/// <param name="showWriteStep">optional object showWriteStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		/// <param name="showDataStep">optional object showDataStep</param>
		/// <param name="showWriteStep">optional object showWriteStep</param>
		/// <param name="showPreviewStep">optional object showPreviewStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
		}

		#endregion

		#pragma warning restore
	}
}


