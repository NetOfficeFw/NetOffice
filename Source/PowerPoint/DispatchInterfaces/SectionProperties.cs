using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface SectionProperties 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743911.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("BA72E551-4FF5-48F4-8215-5505F990966F")]
	public interface SectionProperties : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745766.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744295.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744380.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746414.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		string Name(Int32 sectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745975.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Rename(Int32 sectionIndex, string sectionName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745367.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 SlidesCount(Int32 sectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744059.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 FirstSlide(Int32 sectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745926.aspx </remarks>
		/// <param name="slideIndex">Int32 slideIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 AddBeforeSlide(Int32 slideIndex, string sectionName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">optional object sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 AddSection(Int32 sectionIndex, object sectionName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 AddSection(Int32 sectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746717.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="toPos">Int32 toPos</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Move(Int32 sectionIndex, Int32 toPos);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744948.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="deleteSlides">bool deleteSlides</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Delete(Int32 sectionIndex, bool deleteSlides);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746673.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		string SectionID(Int32 sectionIndex);

		#endregion
	}
}
