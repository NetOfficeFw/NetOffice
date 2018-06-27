using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface ModalBrowser 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ModalBrowser : COMObject, NetOffice.PublisherApi.ModalBrowser
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
                    _contractType = typeof(NetOffice.PublisherApi.ModalBrowser);
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
                    _type = typeof(ModalBrowser);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ModalBrowser() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void TaskCompleted()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TaskCompleted");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lWidth">Int32 lWidth</param>
		/// <param name="lHeight">Int32 lHeight</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ResizeTo(Int32 lWidth, Int32 lHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResizeTo", lWidth, lHeight);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lx">Int32 lx</param>
		/// <param name="ly">Int32 ly</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MoveTo(Int32 lx, Int32 ly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveTo", lx, ly);
		}

		#endregion

		#pragma warning restore
	}
}

