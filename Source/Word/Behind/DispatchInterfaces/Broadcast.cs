using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Broadcast 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229208.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Broadcast : COMObject, NetOffice.WordApi.Broadcast
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
                    _contractType = typeof(NetOffice.WordApi.Broadcast);
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
                    _type = typeof(Broadcast);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Broadcast() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231615.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231440.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228838.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual string AttendeeUrl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AttendeeUrl");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231748.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual NetOffice.OfficeApi.Enums.MsoBroadcastState State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBroadcastState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230812.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual Int32 Capabilities
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Capabilities");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228473.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual string PresenterServiceUrl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PresenterServiceUrl");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230720.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual string SessionID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SessionID");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227544.aspx </remarks>
		/// <param name="serverUrl">string serverUrl</param>
		[SupportByVersion("Word", 15, 16)]
		public virtual void Start(string serverUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Start", serverUrl);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228106.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual void Pause()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Pause");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230520.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual void Resume()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resume");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232060.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual void End()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "End");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232140.aspx </remarks>
		/// <param name="notesUrl">string notesUrl</param>
		/// <param name="notesWacUrl">string notesWacUrl</param>
		[SupportByVersion("Word", 15, 16)]
		public virtual void AddMeetingNotes(string notesUrl, string notesWacUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddMeetingNotes", notesUrl, notesWacUrl);
		}

		#endregion

		#pragma warning restore
	}
}


