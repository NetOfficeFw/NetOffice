using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface Sequence 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744554.aspx </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Sequence : Collection, NetOffice.PowerPointApi.Sequence
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
                    _contractType = typeof(NetOffice.PowerPointApi.Sequence);
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
                    _type = typeof(Sequence);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Sequence() : base()
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
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
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
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
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Item", typeof(NetOffice.PowerPointApi.Effect), index);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", typeof(NetOffice.PowerPointApi.Effect), new object[]{ shape, effectId, level, trigger, index });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", typeof(NetOffice.PowerPointApi.Effect), shape, effectId);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", typeof(NetOffice.PowerPointApi.Effect), shape, effectId, level);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddEffect", typeof(NetOffice.PowerPointApi.Effect), shape, effectId, level, trigger);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Clone", typeof(NetOffice.PowerPointApi.Effect), effect, index);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "Clone", typeof(NetOffice.PowerPointApi.Effect), effect);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744048.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationFor(NetOffice.PowerPointApi.Shape shape)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "FindFirstAnimationFor", typeof(NetOffice.PowerPointApi.Effect), shape);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746508.aspx </remarks>
		/// <param name="click">Int32 click</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationForClick(Int32 click)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "FindFirstAnimationForClick", typeof(NetOffice.PowerPointApi.Effect), click);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToBuildLevel", typeof(NetOffice.PowerPointApi.Effect), effect, level);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", typeof(NetOffice.PowerPointApi.Effect), effect, after, dimColor, dimSchemeColor);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", typeof(NetOffice.PowerPointApi.Effect), effect, after);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAfterEffect", typeof(NetOffice.PowerPointApi.Effect), effect, after, dimColor);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAnimateBackground", typeof(NetOffice.PowerPointApi.Effect), effect, animateBackground);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToAnimateInReverse", typeof(NetOffice.PowerPointApi.Effect), effect, animateInReverse);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "ConvertToTextUnitEffect", typeof(NetOffice.PowerPointApi.Effect), effect, unitEffect);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", typeof(NetOffice.PowerPointApi.Effect), new object[]{ pShape, effectId, trigger, pTriggerShape, bookmark, level });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", typeof(NetOffice.PowerPointApi.Effect), pShape, effectId, trigger, pTriggerShape);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Effect>(this, "AddTriggerEffect", typeof(NetOffice.PowerPointApi.Effect), new object[]{ pShape, effectId, trigger, pTriggerShape, bookmark });
		}

		#endregion

		#pragma warning restore
	}
}


