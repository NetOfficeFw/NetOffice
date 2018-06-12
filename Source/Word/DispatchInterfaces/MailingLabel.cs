using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface MailingLabel 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835169.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface MailingLabel : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837248.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840786.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191949.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool DefaultPrintBarCode { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845366.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdPaperTray DefaultLaserTray { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837913.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.CustomLabels CustomLabels { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string DefaultLabelName { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835161.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool Vertical { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		/// <param name="vertical">optional object vertical</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		/// <param name="vertical">optional object vertical</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000(object name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000(object name, object address);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836933.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void LabelOptions();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		/// <param name="vertical">optional object vertical</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		/// <param name="vertical">optional object vertical</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="laserTray">optional object laserTray</param>
		/// <param name="singleLabel">optional object singleLabel</param>
		/// <param name="row">optional object row</param>
		/// <param name="column">optional object column</param>
		/// <param name="printEPostageLabel">optional object printEPostageLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel);

		#endregion
	}
}
