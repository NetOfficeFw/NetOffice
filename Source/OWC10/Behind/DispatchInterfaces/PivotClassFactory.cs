using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotClassFactory 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotClassFactory : COMObject, NetOffice.OWC10Api.PivotClassFactory
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotClassFactory);
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
                    _type = typeof(PivotClassFactory);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotClassFactory() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell detailCell</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NewDetailCell", detailCell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewDetailCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell detailCell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewDetailCell")]
		public virtual object NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
		{
			return get_NewDetailCell(detailCell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate aggregate</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NewAggregate", aggregate);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewAggregate
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate aggregate</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewAggregate")]
		public virtual object NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
		{
			return get_NewAggregate(aggregate);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember rowMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NewRowMember", rowMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewRowMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember rowMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewRowMember")]
		public virtual object NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
		{
			return get_NewRowMember(rowMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember columnMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NewColumnMember", columnMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewColumnMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember columnMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewColumnMember")]
		public virtual object NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
		{
			return get_NewColumnMember(columnMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell cell</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_NewCell(NetOffice.OWC10Api.PivotCell cell)
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NewCell", cell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell cell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewCell")]
		public virtual object NewCell(NetOffice.OWC10Api.PivotCell cell)
		{
			return get_NewCell(cell);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

