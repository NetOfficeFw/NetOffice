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
 	public class Envelope : COMObject
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
                    _type = typeof(Envelope);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Envelope(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Envelope(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Envelope(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839987.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837451.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837283.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844876.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Address
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Address", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ReturnAddress
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "ReturnAddress", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DefaultPrintBarCode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultPrintBarCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultPrintBarCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DefaultPrintFIMA
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultPrintFIMA");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultPrintFIMA", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195334.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single DefaultHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "DefaultHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single DefaultWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "DefaultWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DefaultSize
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DefaultOmitReturnAddress
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultOmitReturnAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultOmitReturnAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837953.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperTray FeedSource
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "FeedSource");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FeedSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single AddressFromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "AddressFromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddressFromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194512.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single AddressFromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "AddressFromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddressFromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836104.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single ReturnAddressFromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ReturnAddressFromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReturnAddressFromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845802.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single ReturnAddressFromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ReturnAddressFromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReturnAddressFromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194331.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Style AddressStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "AddressStyle", NetOffice.WordApi.Style.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838363.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Style ReturnAddressStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "ReturnAddressStyle", NetOffice.WordApi.Style.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEnvelopeOrientation DefaultOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEnvelopeOrientation>(this, "DefaultOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838355.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DefaultFaceUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultFaceUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultFaceUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192380.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool Vertical
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Vertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Vertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838725.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single RecipientNamefromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RecipientNamefromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientNamefromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196823.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single RecipientNamefromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RecipientNamefromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientNamefromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838472.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single RecipientPostalfromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RecipientPostalfromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientPostalfromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837337.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single RecipientPostalfromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RecipientPostalfromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientPostalfromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194048.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single SenderNamefromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SenderNamefromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderNamefromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844790.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single SenderNamefromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SenderNamefromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderNamefromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194353.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single SenderPostalfromLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SenderPostalfromLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderPostalfromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835494.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single SenderPostalfromTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SenderPostalfromTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderPostalfromTop", value);
			}
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft, senderPostalfromTop });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Insert()
		{
			 Factory.ExecuteMethod(this, "Insert");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Insert(object extractAddress)
		{
			 Factory.ExecuteMethod(this, "Insert", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Insert(object extractAddress, object address)
		{
			 Factory.ExecuteMethod(this, "Insert", extractAddress, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Insert(object extractAddress, object address, object autoText)
		{
			 Factory.ExecuteMethod(this, "Insert", extractAddress, address, autoText);
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 Factory.ExecuteMethod(this, "Insert", extractAddress, address, autoText, omitReturnAddress);
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop });
		}

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
		public void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft, senderPostalfromTop });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object extractAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object extractAddress, object address)
		{
			 Factory.ExecuteMethod(this, "PrintOut", extractAddress, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object extractAddress, object address, object autoText)
		{
			 Factory.ExecuteMethod(this, "PrintOut", extractAddress, address, autoText);
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress);
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop });
		}

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
		public void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197914.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UpdateDocument()
		{
			 Factory.ExecuteMethod(this, "UpdateDocument");
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Insert2000()
		{
			 Factory.ExecuteMethod(this, "Insert2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Insert2000(object extractAddress)
		{
			 Factory.ExecuteMethod(this, "Insert2000", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Insert2000(object extractAddress, object address)
		{
			 Factory.ExecuteMethod(this, "Insert2000", extractAddress, address);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Insert2000(object extractAddress, object address, object autoText)
		{
			 Factory.ExecuteMethod(this, "Insert2000", extractAddress, address, autoText);
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 Factory.ExecuteMethod(this, "Insert2000", extractAddress, address, autoText, omitReturnAddress);
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
		}

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
		public void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 Factory.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			 Factory.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object extractAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object extractAddress, object address)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", extractAddress, address);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object extractAddress, object address, object autoText)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", extractAddress, address, autoText);
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", extractAddress, address, autoText, omitReturnAddress);
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
		}

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
		public void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196101.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Options()
		{
			 Factory.ExecuteMethod(this, "Options");
		}

		#endregion

		#pragma warning restore
	}
}
