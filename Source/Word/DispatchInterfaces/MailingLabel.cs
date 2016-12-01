using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface MailingLabel 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835169.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MailingLabel : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailingLabel(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837248.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840786.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191949.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DefaultPrintBarCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultPrintBarCode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultPrintBarCode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845366.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperTray DefaultLaserTray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultLaserTray", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdPaperTray)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultLaserTray", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837913.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.CustomLabels CustomLabels
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomLabels", paramsArray);
				NetOffice.WordApi.CustomLabels newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.CustomLabels.LateBindingApiWrapperType) as NetOffice.WordApi.CustomLabels;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840714.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string DefaultLabelName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultLabelName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultLabelName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835161.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool Vertical
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Vertical", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Vertical", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress, laserTray);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		/// <param name="vertical">optional object Vertical</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835757.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument(object name, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress, laserTray, printEPostageLabel);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row, column);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		/// <param name="vertical">optional object Vertical</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193415.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress, laserTray);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocument2000(object name, object address, object autoText, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, autoText, extractAddress);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocument2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row, column);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object name, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, address, extractAddress, laserTray, singleLabel, row);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836933.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void LabelOptions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LabelOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		/// <param name="vertical">optional object Vertical</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel, object vertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, autoText, extractAddress, laserTray, printEPostageLabel, vertical);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, autoText);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, autoText, extractAddress);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, autoText, extractAddress, laserTray);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196548.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CreateNewDocumentByID(object labelID, object address, object autoText, object extractAddress, object laserTray, object printEPostageLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, autoText, extractAddress, laserTray, printEPostageLabel);
			object returnItem = Invoker.MethodReturn(this, "CreateNewDocumentByID", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		/// <param name="vertical">optional object Vertical</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel, object vertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel, vertical);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray, singleLabel);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray, singleLabel, row);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray, singleLabel, row, column);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822145.aspx
		/// </summary>
		/// <param name="labelID">optional object LabelID</param>
		/// <param name="address">optional object Address</param>
		/// <param name="extractAddress">optional object ExtractAddress</param>
		/// <param name="laserTray">optional object LaserTray</param>
		/// <param name="singleLabel">optional object SingleLabel</param>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		/// <param name="printEPostageLabel">optional object PrintEPostageLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void PrintOutByID(object labelID, object address, object extractAddress, object laserTray, object singleLabel, object row, object column, object printEPostageLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(labelID, address, extractAddress, laserTray, singleLabel, row, column, printEPostageLabel);
			Invoker.Method(this, "PrintOutByID", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}