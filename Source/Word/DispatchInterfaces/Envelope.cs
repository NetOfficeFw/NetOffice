using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Envelope 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844948.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Envelope : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839987.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837451.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837283.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844876.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Address { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ReturnAddress { get; }

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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool DefaultPrintFIMA { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195334.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single DefaultHeight { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single DefaultWidth { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string DefaultSize { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool DefaultOmitReturnAddress { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837953.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdPaperTray FeedSource { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single AddressFromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194512.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single AddressFromTop { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836104.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single ReturnAddressFromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845802.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single ReturnAddressFromTop { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194331.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Style AddressStyle { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838363.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Style ReturnAddressStyle { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdEnvelopeOrientation DefaultOrientation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838355.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool DefaultFaceUp { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192380.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool Vertical { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838725.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single RecipientNamefromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196823.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single RecipientNamefromTop { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838472.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single RecipientPostalfromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837337.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single RecipientPostalfromTop { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194048.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single SenderNamefromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844790.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single SenderNamefromTop { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194353.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single SenderPostalfromLeft { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835494.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single SenderPostalfromTop { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		/// <param name="senderPostalfromLeft">optional object senderPostalfromLeft</param>
		/// <param name="senderPostalfromTop">optional object senderPostalfromTop</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		/// <param name="senderPostalfromLeft">optional object senderPostalfromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		/// <param name="senderPostalfromLeft">optional object senderPostalfromLeft</param>
		/// <param name="senderPostalfromTop">optional object senderPostalfromTop</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		/// <param name="printEPostage">optional object printEPostage</param>
		/// <param name="vertical">optional object vertical</param>
		/// <param name="recipientNamefromLeft">optional object recipientNamefromLeft</param>
		/// <param name="recipientNamefromTop">optional object recipientNamefromTop</param>
		/// <param name="recipientPostalfromLeft">optional object recipientPostalfromLeft</param>
		/// <param name="recipientPostalfromTop">optional object recipientPostalfromTop</param>
		/// <param name="senderNamefromLeft">optional object senderNamefromLeft</param>
		/// <param name="senderNamefromTop">optional object senderNamefromTop</param>
		/// <param name="senderPostalfromLeft">optional object senderPostalfromLeft</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197914.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UpdateDocument();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		/// <param name="defaultOrientation">optional object defaultOrientation</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation);

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
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="omitReturnAddress">optional object omitReturnAddress</param>
		/// <param name="returnAddress">optional object returnAddress</param>
		/// <param name="returnAutoText">optional object returnAutoText</param>
		/// <param name="printBarCode">optional object printBarCode</param>
		/// <param name="printFIMA">optional object printFIMA</param>
		/// <param name="size">optional object size</param>
		/// <param name="height">optional object height</param>
		/// <param name="width">optional object width</param>
		/// <param name="feedSource">optional object feedSource</param>
		/// <param name="addressFromLeft">optional object addressFromLeft</param>
		/// <param name="addressFromTop">optional object addressFromTop</param>
		/// <param name="returnAddressFromLeft">optional object returnAddressFromLeft</param>
		/// <param name="returnAddressFromTop">optional object returnAddressFromTop</param>
		/// <param name="defaultFaceUp">optional object defaultFaceUp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196101.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Options();

		#endregion
	}
}
