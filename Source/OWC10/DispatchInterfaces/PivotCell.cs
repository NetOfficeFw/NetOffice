using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotCell 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotCell : COMObject
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
                    _type = typeof(PivotCell);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotCell(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotCell(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotCell(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotAggregates Aggregates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotAggregates>(this, "Aggregates", NetOffice.OWC10Api.PivotAggregates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool Expanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Expanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Expanded", value);
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
		public NetOffice.OWC10Api.PivotRowMember RowMember
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotRowMember>(this, "RowMember", NetOffice.OWC10Api.PivotRowMember.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMember ColumnMember
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "ColumnMember", NetOffice.OWC10Api.PivotColumnMember.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailTop", value);
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
		public NetOffice.OWC10Api.PivotDetailCell get_DetailCells(Int32 row, Int32 column)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDetailCell>(this, "DetailCells", NetOffice.OWC10Api.PivotDetailCell.LateBindingApiWrapperType, row, column);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailCells
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailCells")]
		public NetOffice.OWC10Api.PivotDetailCell DetailCells(Int32 row, Int32 column)
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
		public NetOffice.OWC10Api.PivotDetailRange get_DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDetailRange>(this, "DetailRange", NetOffice.OWC10Api.PivotDetailRange.LateBindingApiWrapperType, topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailRange
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotDetailCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotDetailCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailRange")]
		public NetOffice.OWC10Api.PivotDetailRange DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight)
		{
			return get_DetailRange(topLeft, bottomRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotData Data
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotData>(this, "Data", NetOffice.OWC10Api.PivotData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailTopOffset
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailTopOffset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailTopOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailRowCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailRowCount");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailColumnCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailColumnCount");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotPageMember PageMember
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotPageMember>(this, "PageMember", NetOffice.OWC10Api.PivotPageMember.LateBindingApiWrapperType);
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
		public void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset, object update)
		{
			 Factory.ExecuteMethod(this, "MoveDetailTop", detailTop, detailTopOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailTop">Int32 detailTop</param>
		/// <param name="detailTopOffset">Int32 detailTopOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset)
		{
			 Factory.ExecuteMethod(this, "MoveDetailTop", detailTop, detailTopOffset);
		}

		#endregion

		#pragma warning restore
	}
}
