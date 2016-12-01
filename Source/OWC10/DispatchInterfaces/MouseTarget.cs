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
	/// DispatchInterface MouseTarget 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MouseTarget : COMObject
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
                    _type = typeof(MouseTarget);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MouseTarget(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MouseTarget(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 Cursor</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseEnter(Int32 x, Int32 y, Int32 cursor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, cursor);
			Invoker.Method(this, "MouseEnter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 Cursor</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseOver(Int32 x, Int32 y, Int32 cursor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, cursor);
			Invoker.Method(this, "MouseOver", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseLeave()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MouseLeave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 Button</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseDown(Int32 x, Int32 y, Int32 button)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, button);
			Invoker.Method(this, "MouseDown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 Button</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseUp(Int32 x, Int32 y, Int32 button)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, button);
			Invoker.Method(this, "MouseUp", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 Button</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseClick(Int32 x, Int32 y, Int32 button)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, button);
			Invoker.Method(this, "MouseClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 Button</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseDblClick(Int32 x, Int32 y, Int32 button)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, button);
			Invoker.Method(this, "MouseDblClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="page">bool Page</param>
		/// <param name="count">Int32 Count</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MouseWheel(bool page, Int32 count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(page, count);
			Invoker.Method(this, "MouseWheel", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}