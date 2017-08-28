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
 	public class MediaFormat : COMObject
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
                    _type = typeof(MediaFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MediaFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MediaFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(string progId) : base(progId)
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
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
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
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
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
				return Factory.ExecuteSinglePropertyGet(this, "Volume");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Volume", value);
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
				return Factory.ExecuteBoolPropertyGet(this, "Muted");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Muted", value);
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
				return Factory.ExecuteInt32PropertyGet(this, "Length");
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
				return Factory.ExecuteInt32PropertyGet(this, "StartPoint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartPoint", value);
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
				return Factory.ExecuteInt32PropertyGet(this, "EndPoint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EndPoint", value);
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
				return Factory.ExecuteInt32PropertyGet(this, "FadeInDuration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FadeInDuration", value);
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
				return Factory.ExecuteInt32PropertyGet(this, "FadeOutDuration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FadeOutDuration", value);
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
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.MediaBookmarks>(this, "MediaBookmarks", NetOffice.PowerPointApi.MediaBookmarks.LateBindingApiWrapperType);
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
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpMediaTaskStatus>(this, "ResamplingStatus");
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
				return Factory.ExecuteBoolPropertyGet(this, "IsLinked");
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
				return Factory.ExecuteBoolPropertyGet(this, "IsEmbedded");
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
				return Factory.ExecuteInt32PropertyGet(this, "AudioSamplingRate");
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
				return Factory.ExecuteInt32PropertyGet(this, "VideoFrameRate");
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
				return Factory.ExecuteInt32PropertyGet(this, "SampleHeight");
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
				return Factory.ExecuteInt32PropertyGet(this, "SampleWidth");
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
				return Factory.ExecuteStringPropertyGet(this, "VideoCompressionType");
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
				return Factory.ExecuteStringPropertyGet(this, "AudioCompressionType");
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
			 Factory.ExecuteMethod(this, "SetDisplayPicture", position);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746677.aspx </remarks>
		/// <param name="filePath">string filePath</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetDisplayPictureFromFile(string filePath)
		{
			 Factory.ExecuteMethod(this, "SetDisplayPictureFromFile", filePath);
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
			 Factory.ExecuteMethod(this, "Resample", new object[]{ trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate, videoBitRate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746339.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Resample()
		{
			 Factory.ExecuteMethod(this, "Resample");
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
			 Factory.ExecuteMethod(this, "Resample", trim);
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
			 Factory.ExecuteMethod(this, "Resample", trim, sampleHeight);
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
			 Factory.ExecuteMethod(this, "Resample", trim, sampleHeight, sampleWidth);
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
			 Factory.ExecuteMethod(this, "Resample", trim, sampleHeight, sampleWidth, videoFrameRate);
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
			 Factory.ExecuteMethod(this, "Resample", new object[]{ trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		/// <param name="profile">optional NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile = 2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ResampleFromProfile(object profile)
		{
			 Factory.ExecuteMethod(this, "ResampleFromProfile", profile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746706.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ResampleFromProfile()
		{
			 Factory.ExecuteMethod(this, "ResampleFromProfile");
		}

		#endregion

		#pragma warning restore
	}
}
