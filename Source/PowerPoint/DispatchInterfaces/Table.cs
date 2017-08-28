using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Table 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746629.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Table : COMObject
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
                    _type = typeof(Table);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Table(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Table(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745250.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744159.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743913.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Columns Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Columns>(this, "Columns", NetOffice.PowerPointApi.Columns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746738.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Rows Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Rows>(this, "Rows", NetOffice.PowerPointApi.Rows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744662.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpDirection TableDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpDirection>(this, "TableDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TableDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744784.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool FirstRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FirstRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745990.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool LastRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LastRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LastRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744530.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool FirstCol
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FirstCol");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstCol", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746304.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool LastCol
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LastCol");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LastCol", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744939.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool HorizBanding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HorizBanding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HorizBanding", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746483.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool VertBanding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VertBanding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VertBanding", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743893.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.TableStyle Style
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.TableStyle>(this, "Style", NetOffice.PowerPointApi.TableStyle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744088.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.TableBackground Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.TableBackground>(this, "Background", NetOffice.PowerPointApi.TableBackground.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746435.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string AlternativeText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternativeText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternativeText", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746100.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744398.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Cell Cell(Int32 row, Int32 column)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Cell>(this, "Cell", NetOffice.PowerPointApi.Cell.LateBindingApiWrapperType, row, column);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void MergeBorders()
		{
			 Factory.ExecuteMethod(this, "MergeBorders");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744157.aspx </remarks>
		/// <param name="scale">Single scale</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ScaleProportionally(Single scale)
		{
			 Factory.ExecuteMethod(this, "ScaleProportionally", scale);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744636.aspx </remarks>
		/// <param name="styleID">optional string StyleID = </param>
		/// <param name="saveFormatting">optional bool SaveFormatting = false</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ApplyStyle(object styleID, object saveFormatting)
		{
			 Factory.ExecuteMethod(this, "ApplyStyle", styleID, saveFormatting);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744636.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ApplyStyle()
		{
			 Factory.ExecuteMethod(this, "ApplyStyle");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744636.aspx </remarks>
		/// <param name="styleID">optional string StyleID = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ApplyStyle(object styleID)
		{
			 Factory.ExecuteMethod(this, "ApplyStyle", styleID);
		}

		#endregion

		#pragma warning restore
	}
}
