using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface MouseTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MouseTarget : COMObject
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
                    _type = typeof(MouseTarget);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MouseTarget(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 cursor</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseEnter(Int32 x, Int32 y, Int32 cursor)
		{
			 Factory.ExecuteMethod(this, "MouseEnter", x, y, cursor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 cursor</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseOver(Int32 x, Int32 y, Int32 cursor)
		{
			 Factory.ExecuteMethod(this, "MouseOver", x, y, cursor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void MouseLeave()
		{
			 Factory.ExecuteMethod(this, "MouseLeave");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseDown(Int32 x, Int32 y, Int32 button)
		{
			 Factory.ExecuteMethod(this, "MouseDown", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseUp(Int32 x, Int32 y, Int32 button)
		{
			 Factory.ExecuteMethod(this, "MouseUp", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseClick(Int32 x, Int32 y, Int32 button)
		{
			 Factory.ExecuteMethod(this, "MouseClick", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseDblClick(Int32 x, Int32 y, Int32 button)
		{
			 Factory.ExecuteMethod(this, "MouseDblClick", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="page">bool page</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("OWC10", 1)]
		public void MouseWheel(bool page, Int32 count)
		{
			 Factory.ExecuteMethod(this, "MouseWheel", page, count);
		}

		#endregion

		#pragma warning restore
	}
}
