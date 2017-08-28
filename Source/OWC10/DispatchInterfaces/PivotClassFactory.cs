using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotClassFactory 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotClassFactory : COMObject
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
                    _type = typeof(PivotClassFactory);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotClassFactory(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotClassFactory(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotClassFactory(string progId) : base(progId)
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
		public object get_NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
		{
			return Factory.ExecuteReferencePropertyGet(this, "NewDetailCell", detailCell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewDetailCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell detailCell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewDetailCell")]
		public object NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
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
		public object get_NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
		{
			return Factory.ExecuteReferencePropertyGet(this, "NewAggregate", aggregate);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewAggregate
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate aggregate</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewAggregate")]
		public object NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
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
		public object get_NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
		{
			return Factory.ExecuteReferencePropertyGet(this, "NewRowMember", rowMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewRowMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember rowMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewRowMember")]
		public object NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
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
		public object get_NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
		{
			return Factory.ExecuteReferencePropertyGet(this, "NewColumnMember", columnMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewColumnMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember columnMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewColumnMember")]
		public object NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
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
		public object get_NewCell(NetOffice.OWC10Api.PivotCell cell)
		{
			return Factory.ExecuteReferencePropertyGet(this, "NewCell", cell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell cell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewCell")]
		public object NewCell(NetOffice.OWC10Api.PivotCell cell)
		{
			return get_NewCell(cell);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
