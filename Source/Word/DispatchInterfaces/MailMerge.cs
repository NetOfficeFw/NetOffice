using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836701.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailMerge : COMObject
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
                    _type = typeof(MailMerge);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MailMerge(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(string progId) : base(progId)
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
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837179.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839335.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeMainDocType MainDocumentType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeMainDocType>(this, "MainDocumentType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MainDocumentType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840195.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeState State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845069.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeDestination Destination
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeDestination>(this, "Destination");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Destination", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838518.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeDataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMergeDataSource>(this, "DataSource", NetOffice.WordApi.MailMergeDataSource.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193129.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeFields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMergeFields>(this, "Fields", NetOffice.WordApi.MailMergeFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840472.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ViewMailMergeFieldCodes
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ViewMailMergeFieldCodes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewMailMergeFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192581.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SuppressBlankLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SuppressBlankLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuppressBlankLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841091.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MailAsAttachment
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MailAsAttachment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailAsAttachment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820768.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string MailAddressFieldName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailAddressFieldName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailAddressFieldName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820986.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string MailSubject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailSubject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailSubject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845591.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool HighlightMergeFields
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HighlightMergeFields");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HighlightMergeFields", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192784.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeMailFormat MailFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailMergeMailFormat>(this, "MailFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MailFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192539.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string ShowSendToCustom
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShowSendToCustom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSendToCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 WizardState
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WizardState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WizardState", value);
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection, object linkToSource)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection, linkToSource });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource()
		{
			 Factory.ExecuteMethod(this, "CreateDataSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", name, passwordDocument);
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", name, passwordDocument, writePasswordDocument);
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", name, passwordDocument, writePasswordDocument, headerRecord);
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery });
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement });
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1 });
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
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection)
		{
			 Factory.ExecuteMethod(this, "CreateDataSource", new object[]{ name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection });
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
		public void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			 Factory.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument, writePasswordDocument, headerRecord);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name)
		{
			 Factory.ExecuteMethod(this, "CreateHeaderSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument);
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
		public void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "CreateHeaderSource", name, passwordDocument, writePasswordDocument);
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1 });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive, object subType)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive, subType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", name, format);
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
		public void OpenDataSource(string name, object format, object confirmConversions)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", name, format, confirmConversions);
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", name, format, confirmConversions, readOnly);
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement });
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
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object openExclusive)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, openExclusive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", name, format);
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
		public void OpenHeaderSource(string name, object format, object confirmConversions)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", name, format, confirmConversions);
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", name, format, confirmConversions, readOnly);
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
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
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		/// <param name="pause">optional object pause</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Execute(object pause)
		{
			 Factory.ExecuteMethod(this, "Execute", pause);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Execute()
		{
			 Factory.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835814.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Check()
		{
			 Factory.ExecuteMethod(this, "Check");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192805.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EditDataSource()
		{
			 Factory.ExecuteMethod(this, "EditDataSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838561.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EditHeaderSource()
		{
			 Factory.ExecuteMethod(this, "EditHeaderSource");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845149.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EditMainDocument()
		{
			 Factory.ExecuteMethod(this, "EditMainDocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">string type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UseAddressBook(string type)
		{
			 Factory.ExecuteMethod(this, "UseAddressBook", type);
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", name, format);
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
		public void OpenDataSource2000(string name, object format, object confirmConversions)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", name, format, confirmConversions);
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", name, format, confirmConversions, readOnly);
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection });
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
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource2000", new object[]{ name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement });
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", name, format);
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", name, format, confirmConversions);
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", name, format, confirmConversions, readOnly);
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles });
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
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
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "OpenHeaderSource2000", new object[]{ name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
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
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", initialState);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep);
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
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep, showTemplateStep);
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
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", initialState, showDocumentStep, showTemplateStep, showDataStep);
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
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep });
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
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", new object[]{ initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
		}

		#endregion

		#pragma warning restore
	}
}
