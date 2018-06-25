using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Fields 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("7742D36C-49D5-11D3-B65C-00C04F8EF32D")]
	public interface Fields : ICOMObject , NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Field>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Field this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Unlink();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddHorizontalInVertical(NetOffice.PublisherApi.TextRange range, string text);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional object FontSize = 10</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName, object fontSize);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Field>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.PublisherApi.Field> GetEnumerator();

        #endregion
    }
}
