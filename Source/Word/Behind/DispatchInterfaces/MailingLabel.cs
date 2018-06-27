using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface MailingLabel 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835169.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailingLabel : COMObject, NetOffice.WordApi.MailingLabel
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
                    _contractType = typeof(NetOffice.WordApi.MailingLabel);
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
                    _type = typeof(MailingLabel);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailingLabel() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837248.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840786.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191949.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845366.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdPaperTray DefaultLaserTray
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "DefaultLaserTray");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultLaserTray", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837913.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.CustomLabels CustomLabels
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CustomLabels>(this, "CustomLabels", typeof(NetOffice.WordApi.CustomLabels));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string DefaultLabelName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultLabelName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultLabelName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835161.aspx </remarks>
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
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), new object[]{ name, address, autoText, extractAddress, laserTray });
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), new object[]{ name, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), name, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), name, address, autoText);
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), name, address, autoText, extractAddress);
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", typeof(NetOffice.WordApi.Document), new object[]{ name, address, autoText, extractAddress, laserTray, printEPostageLabel });
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column });
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object name, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", name, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object name, object address, object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", name, address, extractAddress);
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", name, address, extractAddress, laserTray);
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel });
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row });
		}

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
		public virtual void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel });
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document), new object[]{ name, address, autoText, extractAddress, laserTray });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument2000()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document));
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument2000(object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document), name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument2000(object name, object address)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document), name, address);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document), name, address, autoText);
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", typeof(NetOffice.WordApi.Document), name, address, autoText, extractAddress);
		}

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
		public virtual void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column });
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
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object name, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", name, address);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object name, object address, object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", name, address, extractAddress);
		}

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
		public virtual void PrintOut2000(object name, object address, object extractAddress, object laserTray)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", name, address, extractAddress, laserTray);
		}

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
		public virtual void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel });
		}

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
		public virtual void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel, row });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836933.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void LabelOptions()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LabelOptions");
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), new object[]{ labelID, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document));
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), labelID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), labelID, address);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), labelID, address, autoText);
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), labelID, address, autoText, extractAddress);
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), new object[]{ labelID, address, autoText, extractAddress, laserTray });
		}

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
		public virtual NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", typeof(NetOffice.WordApi.Document), new object[]{ labelID, address, autoText, extractAddress, laserTray, printEPostageLabel });
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void PrintOutByID()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void PrintOutByID(object labelID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", labelID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void PrintOutByID(object labelID, object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", labelID, address);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		/// <param name="extractAddress">optional object extractAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void PrintOutByID(object labelID, object address, object extractAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", labelID, address, extractAddress);
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", labelID, address, extractAddress, laserTray);
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel });
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row });
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column });
		}

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
		public virtual void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel });
		}

		#endregion

		#pragma warning restore
	}
}


