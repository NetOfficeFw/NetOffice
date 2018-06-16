using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface MediaFormat 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744263.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("BA72E550-4FF5-48F4-8215-5505F990966F")]
	public interface MediaFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744541.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745175.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746131.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Single Volume { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744385.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool Muted { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746068.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Length { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745838.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 StartPoint { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746105.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 EndPoint { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745782.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 FadeInDuration { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746771.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 FadeOutDuration { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746520.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.MediaBookmarks MediaBookmarks { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744315.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Enums.PpMediaTaskStatus ResamplingStatus { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745895.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool IsLinked { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746271.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool IsEmbedded { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744842.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 AudioSamplingRate { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746132.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 VideoFrameRate { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744903.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 SampleHeight { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744690.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 SampleWidth { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744226.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		string VideoCompressionType { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745256.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		string AudioCompressionType { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745841.aspx </remarks>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void SetDisplayPicture(Int32 position);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746677.aspx </remarks>
		/// <param name="filePath">string filePath</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void SetDisplayPictureFromFile(string filePath);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		/// <param name="audioSamplingRate">optional Int32 AudioSamplingRate = 48000</param>
		/// <param name="videoBitRate">optional Int32 VideoBitRate = 7000000</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate, object videoBitRate);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim, object sampleHeight);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim, object sampleHeight, object sampleWidth);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		/// <param name="audioSamplingRate">optional Int32 AudioSamplingRate = 48000</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		/// <param name="profile">optional NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile = 2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void ResampleFromProfile(object profile);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void ResampleFromProfile();

		#endregion
	}
}
