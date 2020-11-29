using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _FormatCondition 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _FormatCondition : COMObject
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
                    _type = typeof(_FormatCondition);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _FormatCondition(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _FormatCondition(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormatCondition(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.ForeColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.BackColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.FontBold"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FontBold
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.FontItalic"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.FontUnderline"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Enabled"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Type"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcFormatConditionType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFormatConditionType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Operator"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFormatConditionOperator>(this, "Operator");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Expression1"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Expression1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Expression1");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Expression2"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Expression2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Expression2");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.ShortestBarLimit"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcFormatBarLimits ShortestBarLimit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFormatBarLimits>(this, "ShortestBarLimit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShortestBarLimit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.ShortestBarValue"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string ShortestBarValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShortestBarValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortestBarValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.LongestBarLimit"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcFormatBarLimits LongestBarLimit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFormatBarLimits>(this, "LongestBarLimit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LongestBarLimit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.LongestBarValue"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string LongestBarValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LongestBarValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LongestBarValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.ShowBarOnly"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool ShowBarOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowBarOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowBarOnly", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Modify"/> </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		/// <param name="expression1">optional object expression1</param>
		/// <param name="expression2">optional object expression2</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Modify(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator, object expression1, object expression2)
		{
			 Factory.ExecuteMethod(this, "Modify", type, _operator, expression1, expression2);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Modify"/> </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Modify(NetOffice.AccessApi.Enums.AcFormatConditionType type)
		{
			 Factory.ExecuteMethod(this, "Modify", type);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Modify"/> </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Modify(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator)
		{
			 Factory.ExecuteMethod(this, "Modify", type, _operator);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Modify"/> </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		/// <param name="expression1">optional object expression1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Modify(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator, object expression1)
		{
			 Factory.ExecuteMethod(this, "Modify", type, _operator, expression1);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.FormatCondition.Delete"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
