using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// Interface IXPath 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IXPath : COMObject, NetOffice.ExcelApi.IXPath
	{
		#pragma warning disable

		#region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IXPath);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IXPath() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlMap Map
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlMap>(this, "Map", typeof(NetOffice.ExcelApi.XmlMap));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool Repeating
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Repeating");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		/// <param name="repeating">optional object repeating</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace, object repeating)
		{
			return Factory.ExecuteInt32MethodGet(this, "SetValue", map, xPath, selectionNamespace, repeating);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath)
		{
			return Factory.ExecuteInt32MethodGet(this, "SetValue", map, xPath);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace)
		{
			return Factory.ExecuteInt32MethodGet(this, "SetValue", map, xPath, selectionNamespace);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 Clear()
		{
			return Factory.ExecuteInt32MethodGet(this, "Clear");
		}

		#endregion

		#pragma warning restore
	}
}


