using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OLEFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OLEFormat : COMObject
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
                    _type = typeof(OLEFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OLEFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OLEFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ClassType"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ClassType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ClassType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ClassType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.DisplayAsIcon"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayAsIcon
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAsIcon");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAsIcon", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.IconName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string IconName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IconName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IconName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.IconPath"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string IconPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IconPath");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.IconIndex"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 IconIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "IconIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IconIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.IconLabel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string IconLabel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IconLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IconLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Label"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Label
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Label");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Object"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ProgID"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ProgID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProgID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.PreserveFormattingOnUpdate"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PreserveFormattingOnUpdate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PreserveFormattingOnUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreserveFormattingOnUpdate", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Activate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Edit"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Edit()
		{
			 Factory.ExecuteMethod(this, "Edit");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.Open"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.DoVerb"/> </remarks>
		/// <param name="verbIndex">optional object verbIndex</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DoVerb(object verbIndex)
		{
			 Factory.ExecuteMethod(this, "DoVerb", verbIndex);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.DoVerb"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DoVerb()
		{
			 Factory.ExecuteMethod(this, "DoVerb");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			 Factory.ExecuteMethod(this, "ConvertTo", new object[]{ classType, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo()
		{
			 Factory.ExecuteMethod(this, "ConvertTo");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo(object classType)
		{
			 Factory.ExecuteMethod(this, "ConvertTo", classType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo(object classType, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "ConvertTo", classType, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName)
		{
			 Factory.ExecuteMethod(this, "ConvertTo", classType, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ConvertTo"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName, object iconIndex)
		{
			 Factory.ExecuteMethod(this, "ConvertTo", classType, displayAsIcon, iconFileName, iconIndex);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OLEFormat.ActivateAs"/> </remarks>
		/// <param name="classType">string classType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ActivateAs(string classType)
		{
			 Factory.ExecuteMethod(this, "ActivateAs", classType);
		}

		#endregion

		#pragma warning restore
	}
}
