﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ErrorCheckingOptions 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ErrorCheckingOptions : COMObject
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
                    _type = typeof(ErrorCheckingOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ErrorCheckingOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ErrorCheckingOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ErrorCheckingOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.Application"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.Creator"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.Parent"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.BackgroundChecking"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool BackgroundChecking
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundChecking");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundChecking", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.IndicatorColorIndex"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlColorIndex IndicatorColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlColorIndex>(this, "IndicatorColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IndicatorColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.EvaluateToError"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EvaluateToError
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EvaluateToError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EvaluateToError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.TextDate"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool TextDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.NumberAsText"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool NumberAsText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NumberAsText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberAsText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.InconsistentFormula"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool InconsistentFormula
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InconsistentFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InconsistentFormula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.OmittedCells"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool OmittedCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OmittedCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OmittedCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.UnlockedFormulaCells"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool UnlockedFormulaCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UnlockedFormulaCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UnlockedFormulaCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.EmptyCellReferences"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EmptyCellReferences
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmptyCellReferences");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmptyCellReferences", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.ListDataValidation"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool ListDataValidation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ListDataValidation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListDataValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ErrorCheckingOptions.InconsistentTableFormula"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool InconsistentTableFormula
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InconsistentTableFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InconsistentTableFormula", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
