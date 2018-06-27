using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotCell 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotCell : COMObject, NetOffice.OWC10Api.PivotCell
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotCell);
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
                    _type = typeof(PivotCell);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotCell() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotAggregates Aggregates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotAggregates>(this, "Aggregates", typeof(NetOffice.OWC10Api.PivotAggregates));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool Expanded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Expanded");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Expanded", value);
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
		public virtual NetOffice.OWC10Api.PivotRowMember RowMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMember>(this, "RowMember", typeof(NetOffice.OWC10Api.PivotRowMember));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotColumnMember ColumnMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "ColumnMember", typeof(NetOffice.OWC10Api.PivotColumnMember));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.PivotDetailCell get_DetailCells(Int32 row, Int32 column)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDetailCell>(this, "DetailCells", typeof(NetOffice.OWC10Api.PivotDetailCell), row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailCells
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailCells")]
		public virtual NetOffice.OWC10Api.PivotDetailCell DetailCells(Int32 row, Int32 column)
		{
			return get_DetailCells(row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotDetailCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotDetailCell bottomRight</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.PivotDetailRange get_DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDetailRange>(this, "DetailRange", typeof(NetOffice.OWC10Api.PivotDetailRange), topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailRange
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotDetailCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotDetailCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailRange")]
		public virtual NetOffice.OWC10Api.PivotDetailRange DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight)
		{
			return get_DetailRange(topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotData Data
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotData>(this, "Data", typeof(NetOffice.OWC10Api.PivotData));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailTopOffset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailTopOffset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailTopOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailRowCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailRowCount");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailColumnCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailColumnCount");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotPageMember PageMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotPageMember>(this, "PageMember", typeof(NetOffice.OWC10Api.PivotPageMember));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailTop">Int32 detailTop</param>
		/// <param name="detailTopOffset">Int32 detailTopOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset, object update)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveDetailTop", detailTop, detailTopOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailTop">Int32 detailTop</param>
		/// <param name="detailTopOffset">Int32 detailTopOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveDetailTop", detailTop, detailTopOffset);
		}

		#endregion

		#pragma warning restore
	}
}


