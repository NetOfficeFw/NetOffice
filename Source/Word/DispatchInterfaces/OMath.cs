using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMath 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821988.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OMath : COMObject
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
                    _type = typeof(OMath);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OMath(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OMath(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197593.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821523.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834527.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836854.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845830.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunctions Functions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFunctions>(this, "Functions", NetOffice.WordApi.OMathFunctions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838750.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845257.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath ParentOMath
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "ParentOMath", NetOffice.WordApi.OMath.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840564.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunction ParentFunction
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFunction>(this, "ParentFunction", NetOffice.WordApi.OMathFunction.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834278.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatRow ParentRow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatRow>(this, "ParentRow", NetOffice.WordApi.OMathMatRow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836630.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatCol ParentCol
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatCol>(this, "ParentCol", NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194990.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath ParentArg
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "ParentArg", NetOffice.WordApi.OMath.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197193.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ArgIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ArgIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845824.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 NestingLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NestingLevel");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197896.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ArgSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ArgSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ArgSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836683.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBreaks Breaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathBreaks>(this, "Breaks", NetOffice.WordApi.OMathBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198098.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathJc Justification
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathJc>(this, "Justification");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Justification", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192630.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 AlignPoint
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AlignPoint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlignPoint", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822894.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Linearize()
		{
			 Factory.ExecuteMethod(this, "Linearize");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198270.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void BuildUp()
		{
			 Factory.ExecuteMethod(this, "BuildUp");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822672.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Remove()
		{
			 Factory.ExecuteMethod(this, "Remove");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838974.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToMathText()
		{
			 Factory.ExecuteMethod(this, "ConvertToMathText");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821266.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToNormalText()
		{
			 Factory.ExecuteMethod(this, "ConvertToNormalText");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838693.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToLiteralText()
		{
			 Factory.ExecuteMethod(this, "ConvertToLiteralText");
		}

		#endregion

		#pragma warning restore
	}
}
