using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface MediaFormat 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744263.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MediaFormat : COMObject, NetOffice.PowerPointApi.MediaFormat
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
                    _contractType = typeof(NetOffice.PowerPointApi.MediaFormat);
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
                    _type = typeof(MediaFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MediaFormat() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744541.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745175.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746131.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single Volume
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Volume");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Volume", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744385.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Muted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Muted");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Muted", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746068.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Length
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Length");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745838.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 StartPoint
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StartPoint");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartPoint", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746105.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 EndPoint
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "EndPoint");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EndPoint", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745782.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 FadeInDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FadeInDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FadeInDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746771.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 FadeOutDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FadeOutDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FadeOutDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746520.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.MediaBookmarks MediaBookmarks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.MediaBookmarks>(this, "MediaBookmarks", typeof(NetOffice.PowerPointApi.MediaBookmarks));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744315.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpMediaTaskStatus ResamplingStatus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpMediaTaskStatus>(this, "ResamplingStatus");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745895.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsLinked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsLinked");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746271.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsEmbedded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsEmbedded");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744842.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 AudioSamplingRate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AudioSamplingRate");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746132.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 VideoFrameRate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "VideoFrameRate");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744903.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 SampleHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SampleHeight");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744690.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 SampleWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SampleWidth");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744226.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string VideoCompressionType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "VideoCompressionType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745256.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string AudioCompressionType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AudioCompressionType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745841.aspx </remarks>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetDisplayPicture(Int32 position)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisplayPicture", position);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746677.aspx </remarks>
		/// <param name="filePath">string filePath</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetDisplayPictureFromFile(string filePath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisplayPictureFromFile", filePath);
		}

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
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate, object videoBitRate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", new object[]{ trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate, videoBitRate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Resample()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Resample(object trim)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", trim);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Resample(object trim, object sampleHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", trim, sampleHeight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Resample(object trim, object sampleHeight, object sampleWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", trim, sampleHeight, sampleWidth);
		}

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
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", trim, sampleHeight, sampleWidth, videoFrameRate);
		}

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
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resample", new object[]{ trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		/// <param name="profile">optional NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile = 2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ResampleFromProfile(object profile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResampleFromProfile", profile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ResampleFromProfile()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResampleFromProfile");
		}

		#endregion

		#pragma warning restore
	}
}


