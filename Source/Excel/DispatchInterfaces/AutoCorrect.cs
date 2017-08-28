using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// AutoCorrect
	/// </summary>
	[SyntaxBypass]
 	public class AutoCorrect_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AutoCorrect_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AutoCorrect_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193969.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ReplacementList(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "ReplacementList", index);
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ReplacementList(object index, object value)
		{
			Factory.ExecutePropertySet(this, "ReplacementList", index, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ReplacementList
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193969.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_ReplacementList")]
		public object ReplacementList(object index)
		{
			return get_ReplacementList(index);
		}

		#endregion

		#region Methods

		#endregion

	}

	/// <summary>
	/// DispatchInterface AutoCorrect 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195348.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AutoCorrect : AutoCorrect_
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
                    _type = typeof(AutoCorrect);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AutoCorrect(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AutoCorrect(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840180.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195350.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840411.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195044.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CapitalizeNamesOfDays
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CapitalizeNamesOfDays");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CapitalizeNamesOfDays", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193969.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ReplacementList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ReplacementList");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ReplacementList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841263.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ReplaceText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReplaceText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReplaceText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838419.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TwoInitialCapitals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TwoInitialCapitals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TwoInitialCapitals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197144.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CorrectSentenceCap
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectSentenceCap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectSentenceCap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192998.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CorrectCapsLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectCapsLock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectCapsLock", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197178.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool DisplayAutoCorrectOptions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAutoCorrectOptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAutoCorrectOptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839028.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool AutoExpandListRange
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoExpandListRange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoExpandListRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840728.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AutoFillFormulasInLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFillFormulasInLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFillFormulasInLists", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196318.aspx </remarks>
		/// <param name="what">string what</param>
		/// <param name="replacement">string replacement</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddReplacement(string what, string replacement)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddReplacement", what, replacement);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823158.aspx </remarks>
		/// <param name="what">string what</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DeleteReplacement(string what)
		{
			return Factory.ExecuteVariantMethodGet(this, "DeleteReplacement", what);
		}

		#endregion

		#pragma warning restore
	}
}
