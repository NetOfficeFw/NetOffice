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
	[TypeId("00020920-0000-0000-C000-000000000046")]
	public interface MailMerge : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837890.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837179.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839335.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdMailMergeMainDocType MainDocumentType { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840195.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdMailMergeState State { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845069.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdMailMergeDestination Destination { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838518.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeDataSource DataSource { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193129.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeFields Fields { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840472.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ViewMailMergeFieldCodes { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192581.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SuppressBlankLines { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841091.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MailAsAttachment { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820768.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string MailAddressFieldName { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820986.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string MailSubject { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845591.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool HighlightMergeFields { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192784.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdMailMergeMailFormat MailFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192539.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		string ShowSendToCustom { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Int32 WizardState { get; set; }

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection, object linkToSource);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateDataSource();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateDataSource(object name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateDataSource(object name, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument);

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord);

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery);

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement);

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1);

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
		void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="headerRecord">optional object headerRecord</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument, object headerRecord);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateHeaderSource(string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateHeaderSource(string name, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive, object subType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenDataSource(string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenDataSource(string name, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenDataSource(string name, object format, object confirmConversions);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement);

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
		void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object openExclusive);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenHeaderSource(string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenHeaderSource(string name, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenHeaderSource(string name, object format, object confirmConversions);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

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
		void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		/// <param name="pause">optional object pause</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Execute(object pause);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Execute();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835814.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Check();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192805.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void EditDataSource();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838561.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void EditHeaderSource();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845149.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void EditMainDocument();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">string type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UseAddressBook(string type);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenDataSource2000(string name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenDataSource2000(string name, object format);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenDataSource2000(string name, object format, object confirmConversions);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection);

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
		void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenHeaderSource2000(string name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenHeaderSource2000(string name, object format);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void OpenHeaderSource2000(string name, object format, object confirmConversions);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

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
		void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

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
		void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ShowWizard(object initialState);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ShowWizard(object initialState, object showDocumentStep);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx </remarks>
		/// <param name="initialState">object initialState</param>
		/// <param name="showDocumentStep">optional object showDocumentStep</param>
		/// <param name="showTemplateStep">optional object showTemplateStep</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep);

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
		void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep);

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
		void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep);

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
		void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep);

		#endregion
	}
}
