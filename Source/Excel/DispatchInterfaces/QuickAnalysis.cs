using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface QuickAnalysis 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231652.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class QuickAnalysis : COMObject
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
                    _type = typeof(QuickAnalysis);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public QuickAnalysis(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public QuickAnalysis(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public QuickAnalysis(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231223.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230558.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228263.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227439.aspx </remarks>
		/// <param name="xlQuickAnalysisMode">optional NetOffice.ExcelApi.Enums.XlQuickAnalysisMode XlQuickAnalysisMode = 0</param>
		[SupportByVersion("Excel", 15, 16)]
		public void Show(object xlQuickAnalysisMode)
		{
			 Factory.ExecuteMethod(this, "Show", xlQuickAnalysisMode);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227439.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public void Show()
		{
			 Factory.ExecuteMethod(this, "Show");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231815.aspx </remarks>
		/// <param name="xlQuickAnalysisMode">optional NetOffice.ExcelApi.Enums.XlQuickAnalysisMode XlQuickAnalysisMode = 0</param>
		[SupportByVersion("Excel", 15, 16)]
		public void Hide(object xlQuickAnalysisMode)
		{
			 Factory.ExecuteMethod(this, "Hide", xlQuickAnalysisMode);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231815.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public void Hide()
		{
			 Factory.ExecuteMethod(this, "Hide");
		}

		#endregion

		#pragma warning restore
	}
}
