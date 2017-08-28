using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotData 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotData : COMObject
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
                    _type = typeof(PivotData);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotData(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotData(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotData(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotView View
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "View", NetOffice.OWC10Api.PivotView.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultRowAxis RowAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultRowAxis>(this, "RowAxis", NetOffice.OWC10Api.PivotResultRowAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultColumnAxis ColumnAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultColumnAxis>(this, "ColumnAxis", NetOffice.OWC10Api.PivotResultColumnAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultFilterAxis FilterAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultFilterAxis>(this, "FilterAxis", NetOffice.OWC10Api.PivotResultFilterAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultDataAxis DataAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultDataAxis>(this, "DataAxis", NetOffice.OWC10Api.PivotResultDataAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMember Left
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "Left", NetOffice.OWC10Api.PivotColumnMember.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotRowMember Top
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMember>(this, "Top", NetOffice.OWC10Api.PivotRowMember.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotCell get_Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "Cells", NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType, row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Cells
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Cells")]
		public NetOffice.OWC10Api.PivotCell Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
		{
			return get_Cells(row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column)
		{
			return Factory.ExecuteInt32PropertyGet(this, "DetailLeft", column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column, Int32 value)
		{
			Factory.ExecutePropertySet(this, "DetailLeft", column, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailLeft
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailLeft")]
		public Int32 DetailLeft(NetOffice.OWC10Api.PivotColumnMember column)
		{
			return get_DetailLeft(column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotCell bottomRight</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotRange get_Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRange>(this, "Range", NetOffice.OWC10Api.PivotRange.LateBindingApiWrapperType, topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public NetOffice.OWC10Api.PivotRange Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight)
		{
			return get_Range(topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Left2
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left2");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Top2
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top2");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultLabel Label
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultLabel>(this, "Label", NetOffice.OWC10Api.PivotResultLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.IPivotControl Control
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.IPivotControl>(this, "Control");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotRowMembers RowMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMembers>(this, "RowMembers", NetOffice.OWC10Api.PivotRowMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotColumnMembers ColumnMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMembers>(this, "ColumnMembers", NetOffice.OWC10Api.PivotColumnMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotCell CurrentCell
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "CurrentCell", NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 LeftOffset
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LeftOffset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 TopOffset
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TopOffset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ViewportTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ViewportTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewportTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ViewportLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ViewportLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewportLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		/// <param name="page">NetOffice.OWC10Api.PivotPageMember page</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotCell get_CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "CellsEx", NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType, row, column, page);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_CellsEx
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		/// <param name="page">NetOffice.OWC10Api.PivotPageMember page</param>
		[SupportByVersion("OWC10", 1), Redirect("get_CellsEx")]
		public NetOffice.OWC10Api.PivotCell CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page)
		{
			return get_CellsEx(row, column, page);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotResultPageAxis PageAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultPageAxis>(this, "PageAxis", NetOffice.OWC10Api.PivotResultPageAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Recordset
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Recordset>(this, "Recordset", NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsConsistent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsConsistent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="top">NetOffice.OWC10Api.PivotRowMember top</param>
		/// <param name="topOffset">Int32 topOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		public void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset, object update)
		{
			 Factory.ExecuteMethod(this, "MoveTop", top, topOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="top">NetOffice.OWC10Api.PivotRowMember top</param>
		/// <param name="topOffset">Int32 topOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset)
		{
			 Factory.ExecuteMethod(this, "MoveTop", top, topOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		public void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset, object update)
		{
			 Factory.ExecuteMethod(this, "MoveLeft", left, leftOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset)
		{
			 Factory.ExecuteMethod(this, "MoveLeft", left, leftOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ShowDetails()
		{
			 Factory.ExecuteMethod(this, "ShowDetails");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void HideDetails()
		{
			 Factory.ExecuteMethod(this, "HideDetails");
		}

		#endregion

		#pragma warning restore
	}
}
