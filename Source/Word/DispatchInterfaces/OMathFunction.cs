using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMathFunction 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192356.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OMathFunction : COMObject
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
                    _type = typeof(OMathFunction);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OMathFunction(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OMathFunction(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822968.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathFunctionType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathFunctionType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821930.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837512.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191735.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821818.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821319.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathArgs Args
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathArgs>(this, "Args", NetOffice.WordApi.OMathArgs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844843.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathAcc Acc
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathAcc>(this, "Acc", NetOffice.WordApi.OMathAcc.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196619.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBar Bar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathBar>(this, "Bar", NetOffice.WordApi.OMathBar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845538.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBox Box
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathBox>(this, "Box", NetOffice.WordApi.OMathBox.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844775.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBorderBox BorderBox
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathBorderBox>(this, "BorderBox", NetOffice.WordApi.OMathBorderBox.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194269.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathDelim Delim
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathDelim>(this, "Delim", NetOffice.WordApi.OMathDelim.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840498.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathEqArray EqArray
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathEqArray>(this, "EqArray", NetOffice.WordApi.OMathEqArray.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840155.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFrac Frac
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFrac>(this, "Frac", NetOffice.WordApi.OMathFrac.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821543.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunc Func
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFunc>(this, "Func", NetOffice.WordApi.OMathFunc.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821338.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathGroupChar GroupChar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathGroupChar>(this, "GroupChar", NetOffice.WordApi.OMathGroupChar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194859.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathLimLow LimLow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathLimLow>(this, "LimLow", NetOffice.WordApi.OMathLimLow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838909.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathLimUpp LimUpp
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathLimUpp>(this, "LimUpp", NetOffice.WordApi.OMathLimUpp.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822992.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMat Mat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMat>(this, "Mat", NetOffice.WordApi.OMathMat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193417.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathNary Nary
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathNary>(this, "Nary", NetOffice.WordApi.OMathNary.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193856.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathPhantom Phantom
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathPhantom>(this, "Phantom", NetOffice.WordApi.OMathPhantom.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840596.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrPre ScrPre
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathScrPre>(this, "ScrPre", NetOffice.WordApi.OMathScrPre.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836951.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathRad Rad
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathRad>(this, "Rad", NetOffice.WordApi.OMathRad.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840167.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSub ScrSub
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathScrSub>(this, "ScrSub", NetOffice.WordApi.OMathScrSub.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821540.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSubSup ScrSubSup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathScrSubSup>(this, "ScrSubSup", NetOffice.WordApi.OMathScrSubSup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840857.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSup ScrSup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathScrSup>(this, "ScrSup", NetOffice.WordApi.OMathScrSup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839640.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath OMath
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "OMath", NetOffice.WordApi.OMath.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196628.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunction Remove()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.OMathFunction>(this, "Remove", NetOffice.WordApi.OMathFunction.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
