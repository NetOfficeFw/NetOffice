using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface DropTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DropTarget : COMObject
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
                    _type = typeof(DropTarget);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DropTarget(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DropTarget(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropTarget(string progId) : base(progId)
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
		public void DragEnter(Int32 x, Int32 y, Int32 keyState, Int32 effect, object _object)
		{
			 Factory.ExecuteMethod(this, "DragEnter", new object[]{ x, y, keyState, effect, _object });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		public void DragOver(Int32 x, Int32 y, Int32 keyState, Int32 effect)
		{
			 Factory.ExecuteMethod(this, "DragOver", x, y, keyState, effect);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void DragLeave()
		{
			 Factory.ExecuteMethod(this, "DragLeave");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		public void Drop(Int32 x, Int32 y, Int32 keyState, Int32 effect)
		{
			 Factory.ExecuteMethod(this, "Drop", x, y, keyState, effect);
		}

		#endregion

		#pragma warning restore
	}
}
