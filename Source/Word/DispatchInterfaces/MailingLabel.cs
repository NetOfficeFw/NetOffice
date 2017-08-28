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
 	public class MailingLabel : COMObject
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
                    _type = typeof(MailingLabel);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MailingLabel(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailingLabel(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(string progId) : base(progId)
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840786.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191949.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845366.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperTray DefaultLaserTray
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "DefaultLaserTray");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultLaserTray", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837913.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.CustomLabels CustomLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CustomLabels>(this, "CustomLabels", NetOffice.WordApi.CustomLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DefaultLabelName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultLabelName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultLabelName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835161.aspx </remarks>
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
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ name, address, autoText, extractAddress, laserTray });
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
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ name, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address);
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
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address, autoText);
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
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address, autoText, extractAddress);
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
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ name, address, autoText, extractAddress, laserTray, printEPostageLabel });
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column });
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name)
		{
			 Factory.ExecuteMethod(this, "PrintOut", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address)
		{
			 Factory.ExecuteMethod(this, "PrintOut", name, address);
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
		public void PrintOut(object name, object address, object extractAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut", name, address, extractAddress);
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray)
		{
			 Factory.ExecuteMethod(this, "PrintOut", name, address, extractAddress, laserTray);
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel });
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row });
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
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel });
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
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ name, address, autoText, extractAddress, laserTray });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address);
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
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address, autoText);
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
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocument2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, name, address, autoText, extractAddress);
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
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel, row, column });
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
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", name, address);
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
		public void PrintOut2000(object name, object address, object extractAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", name, address, extractAddress);
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
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", name, address, extractAddress, laserTray);
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
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel });
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
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ name, address, extractAddress, laserTray, singleLabel, row });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836933.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void LabelOptions()
		{
			 Factory.ExecuteMethod(this, "LabelOptions");
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
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ labelID, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, labelID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, labelID, address);
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
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, labelID, address, autoText);
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
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, labelID, address, autoText, extractAddress);
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
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ labelID, address, autoText, extractAddress, laserTray });
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
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CreateNewDocumentByID", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ labelID, address, autoText, extractAddress, laserTray, printEPostageLabel });
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void PrintOutByID()
		{
			 Factory.ExecuteMethod(this, "PrintOutByID");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", labelID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx </remarks>
		/// <param name="labelID">optional object labelID</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", labelID, address);
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
		public void PrintOutByID(object labelID, object address, object extractAddress)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", labelID, address, extractAddress);
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", labelID, address, extractAddress, laserTray);
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel });
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row });
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column });
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
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			 Factory.ExecuteMethod(this, "PrintOutByID", new object[]{ labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel });
		}

		#endregion

		#pragma warning restore
	}
}
