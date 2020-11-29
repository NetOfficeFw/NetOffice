using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Broadcast 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast"/> </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Broadcast : COMObject
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
                    _type = typeof(Broadcast);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Broadcast(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Broadcast(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Broadcast(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.application"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.parent"/> </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.attendeeurl"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public string AttendeeUrl
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AttendeeUrl");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.state"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.OfficeApi.Enums.MsoBroadcastState State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBroadcastState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.capabilities"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public Int32 Capabilities
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Capabilities");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.presenterserviceurl"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public string PresenterServiceUrl
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PresenterServiceUrl");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.sessionid"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public string SessionID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SessionID");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.start"/> </remarks>
		/// <param name="serverUrl">string serverUrl</param>
		[SupportByVersion("Word", 15, 16)]
		public void Start(string serverUrl)
		{
			 Factory.ExecuteMethod(this, "Start", serverUrl);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.pause"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void Pause()
		{
			 Factory.ExecuteMethod(this, "Pause");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.resume"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void Resume()
		{
			 Factory.ExecuteMethod(this, "Resume");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.end"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void End()
		{
			 Factory.ExecuteMethod(this, "End");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.broadcast.addmeetingnotes"/> </remarks>
		/// <param name="notesUrl">string notesUrl</param>
		/// <param name="notesWacUrl">string notesWacUrl</param>
		[SupportByVersion("Word", 15, 16)]
		public void AddMeetingNotes(string notesUrl, string notesWacUrl)
		{
			 Factory.ExecuteMethod(this, "AddMeetingNotes", notesUrl, notesWacUrl);
		}

		#endregion

		#pragma warning restore
	}
}
