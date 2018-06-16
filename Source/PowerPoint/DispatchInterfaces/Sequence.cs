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
	[TypeId("914934DE-5A91-11CF-8700-00AA0060263B")]
	public interface Sequence : Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745955.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746840.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Effect this[Int32 index] { get; }

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
		NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger, object index);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level);

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
		NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect, object index);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744048.aspx </remarks>
		/// <param name="shape">NetOffice.PowerPointApi.Shape shape</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect FindFirstAnimationFor(NetOffice.PowerPointApi.Shape shape);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746508.aspx </remarks>
		/// <param name="click">Int32 click</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect FindFirstAnimationForClick(Int32 click);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746657.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="level">NetOffice.PowerPointApi.Enums.MsoAnimateByLevel level</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToBuildLevel(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimateByLevel level);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		/// <param name="dimSchemeColor">optional NetOffice.PowerPointApi.Enums.PpColorSchemeIndex DimSchemeColor = 0</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor, object dimSchemeColor);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745293.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="animateBackground">NetOffice.OfficeApi.Enums.MsoTriState animateBackground</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToAnimateBackground(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateBackground);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746429.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="animateInReverse">NetOffice.OfficeApi.Enums.MsoTriState animateInReverse</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToAnimateInReverse(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateInReverse);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746736.aspx </remarks>
		/// <param name="effect">NetOffice.PowerPointApi.Effect effect</param>
		/// <param name="unitEffect">NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Effect ConvertToTextUnitEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect);

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
		NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark, object level);

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
		NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape);

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
		NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark);

		#endregion
	}
}
