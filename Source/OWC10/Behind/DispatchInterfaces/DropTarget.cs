using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface DropTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DropTarget : COMObject, NetOffice.OWC10Api.DropTarget
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
                    _contractType = typeof(NetOffice.OWC10Api.DropTarget);
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
                    _type = typeof(DropTarget);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DropTarget() : base()
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
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		/// <param name="_object">object object</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DragEnter(Int32 x, Int32 y, Int32 keyState, Int32 effect, object _object)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DragEnter", new object[]{ x, y, keyState, effect, _object });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DragOver(Int32 x, Int32 y, Int32 keyState, Int32 effect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DragOver", x, y, keyState, effect);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void DragLeave()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DragLeave");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Drop(Int32 x, Int32 y, Int32 keyState, Int32 effect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Drop", x, y, keyState, effect);
		}

		#endregion

		#pragma warning restore
	}
}

