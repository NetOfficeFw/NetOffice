using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMathMat 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194327.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OMathMat : COMObject
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
                    _type = typeof(OMathMat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OMathMat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OMathMat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839796.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836965.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835716.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821536.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatRows Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatRows>(this, "Rows", NetOffice.WordApi.OMathMatRows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197927.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatCols Cols
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatCols>(this, "Cols", NetOffice.WordApi.OMathMatCols.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196512.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="col">Int32 col</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.OMath get_Cell(Int32 row, Int32 col)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "Cell", NetOffice.WordApi.OMath.LateBindingApiWrapperType, row, col);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Alias for get_Cell
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196512.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="col">Int32 col</param>
		[SupportByVersion("Word", 12,14,15,16), Redirect("get_Cell")]
		public NetOffice.WordApi.OMath Cell(Int32 row, Int32 col)
		{
			return get_Cell(row, col);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845375.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathVertAlignType Align
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathVertAlignType>(this, "Align");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Align", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839960.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool PlcHoldHidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PlcHoldHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PlcHoldHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839195.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathSpacingRule RowSpacingRule
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathSpacingRule>(this, "RowSpacingRule");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RowSpacingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838755.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 RowSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RowSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836905.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ColSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194023.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathSpacingRule ColGapRule
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathSpacingRule>(this, "ColGapRule");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ColGapRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845805.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ColGap
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColGap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColGap", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
