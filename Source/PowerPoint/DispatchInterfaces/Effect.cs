using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// Represents timing information about a slide animation.
	/// Use the <see cref="Sequence.AddEffect"/> method to add an effect. 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Effect : COMObject
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
                    _type = typeof(Effect);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Effect(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Effect(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Effect(string progId) : base(progId)
		{
		}

		#endregion

		#region Properties

		/// <summary>
		/// Returns an Application object that represents the creator of the specified object.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// Returns the parent object for the specified object.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// Returns a Shape object that represents an animated shape.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "Shape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Shape", value);
			}
		}

		/// <summary>
		/// Returns an integer that represents the index number for an animation effect or design. Read-only.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// Returns a Timing object that represents the timing properties for an animation sequence.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Timing Timing
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Timing>(this, "Timing", NetOffice.PowerPointApi.Timing.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// Sets or returns an MsoAnimEffect constant that represents an animation effect type. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.MsoAnimEffect EffectType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.MsoAnimEffect>(this, "EffectType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EffectType", value);
			}
		}

		/// <summary>
		/// Returns an EffectParameters object that represents animation effect properties.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.EffectParameters EffectParameters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.EffectParameters>(this, "EffectParameters", NetOffice.PowerPointApi.EffectParameters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// Returns or sets the start of a text range. Read-only.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 TextRangeStart
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextRangeStart");
			}
		}

		/// <summary>
		/// Returns or sets an integer that represents the length of a text range. Read-only.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 TextRangeLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextRangeLength");
			}
		}

		/// <summary>
		/// Returns or sets the paragraph in a text range to which to apply animation effects. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 Paragraph
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Paragraph");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Paragraph", value);
			}
		}

		/// <summary>
		/// Returns the name of an animation effect. Read-only.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string DisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DisplayName");
			}
		}

		/// <summary>
		/// Determines whether the animation effect is an exit effect. Read/write.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Exit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Exit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Exit", value);
			}
		}

		/// <summary>
		/// Returns a specified slide animation behavior as an AnimationBehaviors collection.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.AnimationBehaviors Behaviors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AnimationBehaviors>(this, "Behaviors", NetOffice.PowerPointApi.AnimationBehaviors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// Returns an EffectInformation object that represents information for a specified animation effect.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.EffectInformation EffectInformation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.EffectInformation>(this, "EffectInformation", NetOffice.PowerPointApi.EffectInformation.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="toPos">Int32 toPos</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void MoveTo(Int32 toPos)
		{
			 Factory.ExecuteMethod(this, "MoveTo", toPos);
		}

		/// <summary>
		/// Moves one animation effect to before another animation effect.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void MoveBefore(NetOffice.PowerPointApi.Effect effect)
		{
			 Factory.ExecuteMethod(this, "MoveBefore", effect);
		}

		/// <summary>
		/// Moves one animation effect to after another animation effect.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void MoveAfter(NetOffice.PowerPointApi.Effect effect)
		{
			 Factory.ExecuteMethod(this, "MoveAfter", effect);
		}

		/// <summary>
		/// Deletes the specified Effect object.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		#endregion

		#pragma warning restore
	}
}
