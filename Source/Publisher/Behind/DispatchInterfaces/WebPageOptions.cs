using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface WebPageOptions 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class WebPageOptions : COMObject, NetOffice.PublisherApi.WebPageOptions
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
                    _contractType = typeof(NetOffice.PublisherApi.WebPageOptions);
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
                    _type = typeof(WebPageOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public WebPageOptions() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Description
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Keywords
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Keywords");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Keywords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool IncludePageOnNewWebNavigationBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IncludePageOnNewWebNavigationBars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IncludePageOnNewWebNavigationBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string BackgroundSound
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BackgroundSound");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackgroundSound", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool BackgroundSoundLoopForever
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BackgroundSoundLoopForever");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 BackgroundSoundLoopCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackgroundSoundLoopCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string PublishFileName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PublishFileName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PublishFileName", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="repeatForever">bool repeatForever</param>
		/// <param name="repeatTimes">optional Int32 RepeatTimes = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetBackgroundSoundRepeat(bool repeatForever, object repeatTimes)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetBackgroundSoundRepeat", repeatForever, repeatTimes);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="repeatForever">bool repeatForever</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetBackgroundSoundRepeat(bool repeatForever)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetBackgroundSoundRepeat", repeatForever);
		}

		#endregion

		#pragma warning restore
	}
}


