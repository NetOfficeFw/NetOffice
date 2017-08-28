using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Validation 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821214.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Validation : COMObject
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
                    _type = typeof(Validation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Validation(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Validation(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Validation(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837359.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822112.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839682.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839985.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AlertStyle
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AlertStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835910.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool IgnoreBlank
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreBlank");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreBlank", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193613.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 IMEMode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "IMEMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IMEMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192976.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool InCellDropdown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InCellDropdown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InCellDropdown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840297.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string ErrorMessage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ErrorMessage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ErrorMessage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838389.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string ErrorTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ErrorTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ErrorTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839245.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string InputMessage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InputMessage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InputMessage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834309.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string InputTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InputTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InputTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837441.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Formula1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula1");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836168.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Formula2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula2");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194880.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Operator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Operator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194576.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ShowError
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835258.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ShowInput
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowInput");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowInput", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834305.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Type
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835574.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Value
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Value");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840078.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlDVType type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Add(NetOffice.ExcelApi.Enums.XlDVType type, object alertStyle, object _operator, object formula1, object formula2)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ type, alertStyle, _operator, formula1, formula2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840078.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlDVType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Add(NetOffice.ExcelApi.Enums.XlDVType type)
		{
			 Factory.ExecuteMethod(this, "Add", type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840078.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlDVType type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Add(NetOffice.ExcelApi.Enums.XlDVType type, object alertStyle)
		{
			 Factory.ExecuteMethod(this, "Add", type, alertStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840078.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlDVType type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Add(NetOffice.ExcelApi.Enums.XlDVType type, object alertStyle, object _operator)
		{
			 Factory.ExecuteMethod(this, "Add", type, alertStyle, _operator);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840078.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlDVType type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Add(NetOffice.ExcelApi.Enums.XlDVType type, object alertStyle, object _operator, object formula1)
		{
			 Factory.ExecuteMethod(this, "Add", type, alertStyle, _operator, formula1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820882.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify(object type, object alertStyle, object _operator, object formula1, object formula2)
		{
			 Factory.ExecuteMethod(this, "Modify", new object[]{ type, alertStyle, _operator, formula1, formula2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify()
		{
			 Factory.ExecuteMethod(this, "Modify");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify(object type)
		{
			 Factory.ExecuteMethod(this, "Modify", type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify(object type, object alertStyle)
		{
			 Factory.ExecuteMethod(this, "Modify", type, alertStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify(object type, object alertStyle, object _operator)
		{
			 Factory.ExecuteMethod(this, "Modify", type, alertStyle, _operator);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198204.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="alertStyle">optional object alertStyle</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Modify(object type, object alertStyle, object _operator, object formula1)
		{
			 Factory.ExecuteMethod(this, "Modify", type, alertStyle, _operator, formula1);
		}

		#endregion

		#pragma warning restore
	}
}
