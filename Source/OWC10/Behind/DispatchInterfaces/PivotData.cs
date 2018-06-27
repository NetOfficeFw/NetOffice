using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotData 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotData : COMObject, NetOffice.OWC10Api.PivotData
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OWC10Api.PivotData);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotData() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotView View
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "View", typeof(NetOffice.OWC10Api.PivotView));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultRowAxis RowAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultRowAxis>(this, "RowAxis", typeof(NetOffice.OWC10Api.PivotResultRowAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultColumnAxis ColumnAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultColumnAxis>(this, "ColumnAxis", typeof(NetOffice.OWC10Api.PivotResultColumnAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultFilterAxis FilterAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultFilterAxis>(this, "FilterAxis", typeof(NetOffice.OWC10Api.PivotResultFilterAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultDataAxis DataAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultDataAxis>(this, "DataAxis", typeof(NetOffice.OWC10Api.PivotResultDataAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotColumnMember Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "Left", typeof(NetOffice.OWC10Api.PivotColumnMember));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotRowMember Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMember>(this, "Top", typeof(NetOffice.OWC10Api.PivotRowMember));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Top", value);
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
		public virtual NetOffice.OWC10Api.PivotCell get_Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "Cells", typeof(NetOffice.OWC10Api.PivotCell), row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Cells
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Cells")]
		public virtual NetOffice.OWC10Api.PivotCell Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
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
		public virtual Int32 get_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailLeft", column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column, Int32 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "DetailLeft", column, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailLeft
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailLeft")]
		public virtual Int32 DetailLeft(NetOffice.OWC10Api.PivotColumnMember column)
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
		public virtual NetOffice.OWC10Api.PivotRange get_Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRange>(this, "Range", typeof(NetOffice.OWC10Api.PivotRange), topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public virtual NetOffice.OWC10Api.PivotRange Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight)
		{
			return get_Range(topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Left2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left2");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Top2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top2");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultLabel Label
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultLabel>(this, "Label", typeof(NetOffice.OWC10Api.PivotResultLabel));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api.IPivotControl Control
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.IPivotControl>(this, "Control");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.PivotRowMembers RowMembers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMembers>(this, "RowMembers", typeof(NetOffice.OWC10Api.PivotRowMembers));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.PivotColumnMembers ColumnMembers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMembers>(this, "ColumnMembers", typeof(NetOffice.OWC10Api.PivotColumnMembers));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotCell CurrentCell
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "CurrentCell", typeof(NetOffice.OWC10Api.PivotCell));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 LeftOffset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LeftOffset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 TopOffset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TopOffset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ViewportTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ViewportTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewportTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ViewportLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ViewportLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewportLeft", value);
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
		public virtual NetOffice.OWC10Api.PivotCell get_CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotCell>(this, "CellsEx", typeof(NetOffice.OWC10Api.PivotCell), row, column, page);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_CellsEx
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		/// <param name="page">NetOffice.OWC10Api.PivotPageMember page</param>
		[SupportByVersion("OWC10", 1), Redirect("get_CellsEx")]
		public virtual NetOffice.OWC10Api.PivotCell CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page)
		{
			return get_CellsEx(row, column, page);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotResultPageAxis PageAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotResultPageAxis>(this, "PageAxis", typeof(NetOffice.OWC10Api.PivotResultPageAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.ADODBApi.Recordset Recordset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Recordset>(this, "Recordset", typeof(NetOffice.ADODBApi.Recordset));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsConsistent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsConsistent");
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
		public virtual void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset, object update)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveTop", top, topOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="top">NetOffice.OWC10Api.PivotRowMember top</param>
		/// <param name="topOffset">Int32 topOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveTop", top, topOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset, object update)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLeft", left, leftOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLeft", left, leftOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowDetails()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowDetails");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void HideDetails()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "HideDetails");
		}

		#endregion

		#pragma warning restore
	}
}


