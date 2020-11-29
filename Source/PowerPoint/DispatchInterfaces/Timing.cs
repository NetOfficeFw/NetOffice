using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// Represents timing properties for an animation effect.
	/// DispatchInterface Timing 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Timing : COMObject
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
                    _type = typeof(Timing);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Timing(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Timing(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Timing(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Application"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Parent"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// Returns or sets the length of an animation in seconds. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.timing.duration"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single Duration
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Duration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Duration", value);
			}
		}

		/// <summary>
		/// Represents the trigger that starts an animation. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.TriggerType"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.MsoAnimTriggerType TriggerType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.MsoAnimTriggerType>(this, "TriggerType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TriggerType", value);
			}
		}

		/// <summary>
		/// Sets or returns the delay, in seconds, from when an animation trigger is enabled. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.TriggerDelayTime"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single TriggerDelayTime
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "TriggerDelayTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TriggerDelayTime", value);
			}
		}

		/// <summary>
		/// Sets or returns a Shape object that represents the shape associated with an animation trigger. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.TriggerShape"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape TriggerShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "TriggerShape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "TriggerShape", value);
			}
		}

		/// <summary>
		/// Sets or returns the number of times to repeat an animation. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.RepeatCount"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 RepeatCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RepeatCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatCount", value);
			}
		}

		/// <summary>
		/// Sets or returns how long repeated animations should last, in seconds. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// An animation will stop at the end of its time sequence or the value of the RepeatDuration property, whichever is shorter.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.RepeatDuration"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single RepeatDuration
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RepeatDuration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Speed"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single Speed
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Speed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Speed", value);
			}
		}

		/// <summary>
		/// Returns or sets the percentage of the duration over which a timing acceleration should take place. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// For example, a value of 0.9 means that an acceleration should start slower than the default speed for 90% of the total animation time, with the last 10% of the animation at the default speed.
		/// To slow down an animation at the end, use the Decelerate property.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Accelerate"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single Accelerate
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Accelerate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Accelerate", value);
			}
		}

		/// <summary>
		/// Sets or returns the percentage of the duration over which a timing deceleration should take place. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// For example, a value of 0.9 means that an deceleration should start at the default speed, and then start to slow down after the first ten percent of the animation.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Decelerate"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single Decelerate
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Decelerate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Decelerate", value);
			}
		}

		/// <summary>
		/// Determines whether an effect should play forward and then in reverse, thereby doubling its duration. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// The value of the AutoReverse property can be one of these MsoTriState constants.
		/// msoFalse - The default. The effect does not play forward and then in reverse.
		/// msoTrue	- The effect plays forward and then in reverse.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.AutoReverse"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AutoReverse
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AutoReverse");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoReverse", value);
			}
		}

		/// <summary>
		/// Determines whether an animation should accelerate when it starts. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// The value of the SmoothStart property can be one of these MsoTriState constants.
		/// msoFalse - The default. The animation does not accelerate when it starts.
		/// msoTrue - The animation accelerates when it starts.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.SmoothStart"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState SmoothStart
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SmoothStart");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SmoothStart", value);
			}
		}

		/// <summary>
		/// Determines whether an animation should decelerate as it ends. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// The value of the SmoothEnd property can be one of these MsoTriState constants.
		/// msoFalse - The default. An animation does not decelerate when it ends.
		/// msoTrue - An animation decelerates when it ends.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.SmoothEnd"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState SmoothEnd
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SmoothEnd");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SmoothEnd", value);
			}
		}

		/// <summary>
		/// Represents whether an object returns to its beginning position after an animation has ended. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>
		/// The value of the RewindAtEnd property can be one of these MsoTriState constants.
		/// msoFalse - The object does not return to its beginning position after an animation has ended.
		/// msoTrue - The object returns to its beginning position after an animation has ended.
		/// Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.RewindAtEnd"/>
		/// </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RewindAtEnd
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "RewindAtEnd");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RewindAtEnd", value);
			}
		}

		/// <summary>
		/// Represents whether the animation effect restarts after the effect has started once. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.Restart"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.MsoAnimEffectRestart Restart
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.MsoAnimEffectRestart>(this, "Restart");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Restart", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.BounceEnd"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState BounceEnd
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "BounceEnd");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BounceEnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.BounceEndIntensity"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single BounceEndIntensity
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BounceEndIntensity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BounceEndIntensity", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Timing.triggerBookmark"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string TriggerBookmark
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TriggerBookmark");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TriggerBookmark", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
