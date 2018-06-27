using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface MouseTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MouseTarget : COMObject, NetOffice.OWC10Api.MouseTarget
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
                    _contractType = typeof(NetOffice.OWC10Api.MouseTarget);
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
                    _type = typeof(MouseTarget);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MouseTarget() : base()
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
		public virtual void MouseEnter(Int32 x, Int32 y, Int32 cursor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseEnter", x, y, cursor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 cursor</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseOver(Int32 x, Int32 y, Int32 cursor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseOver", x, y, cursor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseLeave()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseLeave");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseDown(Int32 x, Int32 y, Int32 button)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseDown", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseUp(Int32 x, Int32 y, Int32 button)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseUp", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseClick(Int32 x, Int32 y, Int32 button)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseClick", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseDblClick(Int32 x, Int32 y, Int32 button)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseDblClick", x, y, button);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="page">bool page</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void MouseWheel(bool page, Int32 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MouseWheel", page, count);
		}

		#endregion

		#pragma warning restore
	}
}

