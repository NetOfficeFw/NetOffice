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
	/// CoClass _CustomControlInReport
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(Events._CustomControlInReportEvents_SinkHelper), typeof(Events.DispCustomControlInReportEvents_SinkHelper))]
    [ComEventInterface(typeof(Events._CustomControlInReportEvents), typeof(Events.DispCustomControlInReportEvents))]
    public class _CustomControlInReport : _CustomControl
	{
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events._CustomControlInReportEvents_SinkHelper __CustomControlInReportEvents_SinkHelper;
        private Events.DispCustomControlInReportEvents_SinkHelper _dispCustomControlInReportEvents_SinkHelper;

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
                    _type = typeof(_CustomControlInReport);
                return _type;
            }
        }

        #endregion

		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CustomControlInReport(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{

		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CustomControlInReport(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{

		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomControlInReport(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomControlInReport(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{

		}

		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomControlInReport(ICOMObject replacedObject) : base(replacedObject)
		{

		}

		/// <summary>
        /// Creates a new instance of _CustomControlInReport
        /// </summary>
		public _CustomControlInReport():base("Access._CustomControlInReport")
		{

		}

		/// <summary>
        /// Creates a new instance of _CustomControlInReport
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public _CustomControlInReport(string progId):base(progId)
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
