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
	/// DispatchInterface PivotRange 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PivotRange : COMObject
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
                    _type = typeof(PivotRange);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotRange(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PivotCell TopLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TopLeft", paramsArray);
				NetOffice.OWC10Api.PivotCell newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType) as NetOffice.OWC10Api.PivotCell;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PivotCell BottomRight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BottomRight", paramsArray);
				NetOffice.OWC10Api.PivotCell newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType) as NetOffice.OWC10Api.PivotCell;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PivotRowMembers RowMembers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowMembers", paramsArray);
				NetOffice.OWC10Api.PivotRowMembers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PivotRowMembers.LateBindingApiWrapperType) as NetOffice.OWC10Api.PivotRowMembers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMembers ColumnMembers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnMembers", paramsArray);
				NetOffice.OWC10Api.PivotColumnMembers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PivotColumnMembers.LateBindingApiWrapperType) as NetOffice.OWC10Api.PivotColumnMembers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember Row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember Column</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotCell get_Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(row, column);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.OWC10Api.PivotCell newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PivotCell.LateBindingApiWrapperType) as NetOffice.OWC10Api.PivotCell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Cells
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember Row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember Column</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PivotCell Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column)
		{
			return get_Cells(row, column);
		}

		#endregion

		#region Methods

		#endregion
		#pragma warning restore
	}
}