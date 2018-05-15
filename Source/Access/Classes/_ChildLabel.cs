using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _ChildLabel
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(Events._ChildLabelEvents_SinkHelper), typeof(Events.DispChildLabelEvents_SinkHelper))]
    [ComEventInterface(typeof(Events._ChildLabelEvents), typeof(Events.DispChildLabelEvents))]
    public class _ChildLabel : _Label
	{
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events._ChildLabelEvents_SinkHelper __ChildLabelEvents_SinkHelper;
        private Events.DispChildLabelEvents_SinkHelper _dispChildLabelEvents_SinkHelper;

		#endregion

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

        /// <summary>
        /// Type Cache
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_ChildLabel);
                return _type;
            }
        }

        #endregion

		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ChildLabel(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{

		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ChildLabel(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{

		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ChildLabel(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ChildLabel(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{

		}

		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ChildLabel(ICOMObject replacedObject) : base(replacedObject)
		{

		}

		/// <summary>
        /// Creates a new instance of _ChildLabel
        /// </summary>
		public _ChildLabel():base("Access._ChildLabel")
		{

		}

		/// <summary>
        /// Creates a new instance of _ChildLabel
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public _ChildLabel(string progId):base(progId)
		{

		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		#endregion

		#pragma warning restore
	}
}
