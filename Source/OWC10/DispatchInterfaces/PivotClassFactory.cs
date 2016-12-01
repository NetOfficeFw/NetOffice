using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface PivotClassFactory 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PivotClassFactory : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell DetailCell</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(detailCell);
			object returnItem = Invoker.PropertyGet(this, "NewDetailCell", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewDetailCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell DetailCell</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell)
		{
			return get_NewDetailCell(detailCell);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate Aggregate</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(aggregate);
			object returnItem = Invoker.PropertyGet(this, "NewAggregate", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewAggregate
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate Aggregate</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate)
		{
			return get_NewAggregate(aggregate);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember RowMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowMember);
			object returnItem = Invoker.PropertyGet(this, "NewRowMember", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewRowMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember RowMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember)
		{
			return get_NewRowMember(rowMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember ColumnMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(columnMember);
			object returnItem = Invoker.PropertyGet(this, "NewColumnMember", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewColumnMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember ColumnMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember)
		{
			return get_NewColumnMember(columnMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell Cell</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_NewCell(NetOffice.OWC10Api.PivotCell cell)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell);
			object returnItem = Invoker.PropertyGet(this, "NewCell", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell Cell</param>
		[SupportByVersionAttribute("OWC10", 1)]
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