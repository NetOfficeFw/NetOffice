using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Envelope 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844948.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Envelope : COMObject, NetOffice.WordApi.Envelope
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
                    _contractType = typeof(NetOffice.WordApi.Envelope);
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
                    _type = typeof(Envelope);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Envelope() : base()
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837451.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837283.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844876.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Address
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Address", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ReturnAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "ReturnAddress", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DefaultPrintBarCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultPrintBarCode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultPrintBarCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DefaultPrintFIMA
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultPrintFIMA");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultPrintFIMA", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195334.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single DefaultHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "DefaultHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192764.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single DefaultWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "DefaultWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string DefaultSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DefaultOmitReturnAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultOmitReturnAddress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultOmitReturnAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837953.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdPaperTray FeedSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "FeedSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FeedSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single AddressFromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "AddressFromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddressFromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194512.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single AddressFromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "AddressFromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddressFromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836104.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single ReturnAddressFromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ReturnAddressFromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReturnAddressFromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845802.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single ReturnAddressFromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ReturnAddressFromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReturnAddressFromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194331.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Style AddressStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "AddressStyle", typeof(NetOffice.WordApi.Style));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838363.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Style ReturnAddressStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "ReturnAddressStyle", typeof(NetOffice.WordApi.Style));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdEnvelopeOrientation DefaultOrientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEnvelopeOrientation>(this, "DefaultOrientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838355.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DefaultFaceUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultFaceUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultFaceUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192380.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool Vertical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Vertical");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Vertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838725.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single RecipientNamefromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RecipientNamefromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecipientNamefromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196823.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single RecipientNamefromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RecipientNamefromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecipientNamefromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838472.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single RecipientPostalfromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RecipientPostalfromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecipientPostalfromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837337.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single RecipientPostalfromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RecipientPostalfromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecipientPostalfromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194048.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SenderNamefromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SenderNamefromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SenderNamefromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844790.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SenderNamefromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SenderNamefromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SenderNamefromTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194353.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SenderPostalfromLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SenderPostalfromLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SenderPostalfromLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835494.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SenderPostalfromTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SenderPostalfromTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SenderPostalfromTop", value);
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft, senderPostalfromTop });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Insert()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Insert(object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198190.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Insert(object extractAddress, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", extractAddress, address);
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
		public virtual void Insert(object extractAddress, object address, object autoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", extractAddress, address, autoText);
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", extractAddress, address, autoText, omitReturnAddress);
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop });
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
		public virtual void Insert(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft, object senderPostalfromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft, senderPostalfromTop });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197594.aspx </remarks>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object extractAddress, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", extractAddress, address);
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
		public virtual void PrintOut(object extractAddress, object address, object autoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", extractAddress, address, autoText);
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress);
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop });
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
		public virtual void PrintOut(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation, object printEPostage, object vertical, object recipientNamefromLeft, object recipientNamefromTop, object recipientPostalfromLeft, object recipientPostalfromTop, object senderNamefromLeft, object senderNamefromTop, object senderPostalfromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation, printEPostage, vertical, recipientNamefromLeft, recipientNamefromTop, recipientPostalfromLeft, recipientPostalfromTop, senderNamefromLeft, senderNamefromTop, senderPostalfromLeft });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197914.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UpdateDocument()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateDocument");
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Insert2000()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Insert2000(object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Insert2000(object extractAddress, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", extractAddress, address);
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
		public virtual void Insert2000(object extractAddress, object address, object autoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", extractAddress, address, autoText);
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", extractAddress, address, autoText, omitReturnAddress);
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
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
		public virtual void Insert2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp, object defaultOrientation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", extractAddress);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="extractAddress">optional object extractAddress</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object extractAddress, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", extractAddress, address);
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", extractAddress, address, autoText);
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", extractAddress, address, autoText, omitReturnAddress);
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop });
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
		public virtual void PrintOut2000(object extractAddress, object address, object autoText, object omitReturnAddress, object returnAddress, object returnAutoText, object printBarCode, object printFIMA, object size, object height, object width, object feedSource, object addressFromLeft, object addressFromTop, object returnAddressFromLeft, object returnAddressFromTop, object defaultFaceUp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft, addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196101.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Options()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Options");
		}

		#endregion

		#pragma warning restore
	}
}


