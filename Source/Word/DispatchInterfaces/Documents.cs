using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Documents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840891.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 9,10,11,12,14,15,16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002096C-0000-0000-C000-000000000046")]
	public interface Documents : ICOMObject, IEnumerableProvider<NetOffice.WordApi.Document>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822958.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195113.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838145.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196684.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.Document this[object index] { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges, object originalFormat, object routeDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges, object originalFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document AddOld(object template, object newTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document AddOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="template">optional object template</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document AddOld(object template);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx </remarks>
		/// <param name="noPrompt">optional object noPrompt</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Save(object noPrompt, object originalFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx </remarks>
		/// <param name="noPrompt">optional object noPrompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Save(object noPrompt);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		/// <param name="documentType">optional object documentType</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType, object visible);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Add();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx </remarks>
		/// <param name="template">optional object template</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Add(object template);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Add(object template, object newTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		/// <param name="documentType">optional object documentType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		/// <param name="xMLTransform">optional object xMLTransform</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
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
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198275.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckOut(string fileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839907.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool CanCheckOut(string fileName);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		/// <param name="xMLTransform">optional object xMLTransform</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838738.aspx </remarks>
		/// <param name="providerID">string providerID</param>
		/// <param name="postURL">string postURL</param>
		/// <param name="blogName">string blogName</param>
		/// <param name="postID">optional string PostID = </param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName, object postID);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838738.aspx </remarks>
		/// <param name="providerID">string providerID</param>
		/// <param name="postURL">string postURL</param>
		/// <param name="blogName">string blogName</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName);

        #endregion

        #region IEnumerable<NetOffice.WordApi.Document>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.Document> GetEnumerator();

        #endregion
    }
}
