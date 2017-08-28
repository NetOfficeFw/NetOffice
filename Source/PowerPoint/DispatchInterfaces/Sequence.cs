using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Sequence 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744554.aspx </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Sequence : Collection
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
                    _type = typeof(Sequence);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Sequence(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Sequence(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745955.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746840.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Effect this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Item", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		/// <param name="trigger">optional NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger = 1</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, new object[]{ shape, effectId, level, trigger, index });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, shape, effectId);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, shape, effectId, level);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		/// <param name="trigger">optional NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, shape, effectId, level, trigger);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Clone", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Clone", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744048.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationFor(NetOffice.PowerPointApi.Shape shape)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "FindFirstAnimationFor", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, shape);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746508.aspx </remarks>
		/// <param name="click">Int32 click</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationForClick(Int32 click)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "FindFirstAnimationForClick", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, click);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746657.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="level">NetOffice.PowerPointApi.Enums.MsoAnimateByLevel level</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToBuildLevel(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimateByLevel level)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToBuildLevel", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, level);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		/// <param name="dimSchemeColor">optional NetOffice.PowerPointApi.Enums.PpColorSchemeIndex DimSchemeColor = 0</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor, object dimSchemeColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, after, dimColor, dimSchemeColor);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, after);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, after, dimColor);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745293.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="animateBackground">NetOffice.OfficeApi.Enums.MsoTriState animateBackground</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAnimateBackground(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateBackground)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAnimateBackground", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, animateBackground);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746429.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="animateInReverse">NetOffice.OfficeApi.Enums.MsoTriState animateInReverse</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAnimateInReverse(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateInReverse)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAnimateInReverse", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, animateInReverse);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746736.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="unitEffect">NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToTextUnitEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToTextUnitEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, effect, unitEffect);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx </remarks>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		/// <param name="bookmark">optional string bookmark = </param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark, object level)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, new object[]{ pShape, effectId, trigger, pTriggerShape, bookmark, level });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx </remarks>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, pShape, effectId, trigger, pTriggerShape);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx </remarks>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		/// <param name="bookmark">optional string bookmark = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType, new object[]{ pShape, effectId, trigger, pTriggerShape, bookmark });
		}

		#endregion

		#pragma warning restore
	}
}
