using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface Broadcast 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745427.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Broadcast : COMObject, NetOffice.PowerPointApi.Broadcast
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
                    _contractType = typeof(NetOffice.PowerPointApi.Broadcast);
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746713.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744882.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744687.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string AttendeeUrl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AttendeeUrl");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744210.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsBroadcasting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsBroadcasting");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230343.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.OfficeApi.Enums.MsoBroadcastState State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBroadcastState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229261.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public Int32 Capabilities
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Capabilities");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230327.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public string SessionID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SessionID");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227435.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public string PresenterServiceUrl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PresenterServiceUrl");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746118.aspx </remarks>
		/// <param name="serverUrl">string serverUrl</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Start(string serverUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Start", serverUrl);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746003.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void End()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "End");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230158.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void Pause()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Pause");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229921.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void Resume()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resume");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229798.aspx </remarks>
		/// <param name="notesUrl">string notesUrl</param>
		/// <param name="notesWacUrl">string notesWacUrl</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void AddMeetingNotes(string notesUrl, string notesWacUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddMeetingNotes", notesUrl, notesWacUrl);
		}

		#endregion

		#pragma warning restore
	}
}


