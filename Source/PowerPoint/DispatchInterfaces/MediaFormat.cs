using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface MediaFormat 
	/// SupportByVersion PowerPoint, 14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MediaFormat : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Single Volume
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Volume", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Volume", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public bool Muted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Muted", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Muted", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 Length
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Length", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 StartPoint
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StartPoint", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StartPoint", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 EndPoint
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EndPoint", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EndPoint", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 FadeInDuration
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FadeInDuration", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FadeInDuration", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 FadeOutDuration
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FadeOutDuration", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FadeOutDuration", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.MediaBookmarks MediaBookmarks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaBookmarks", paramsArray);
				NetOffice.PowerPointApi.MediaBookmarks newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.MediaBookmarks.LateBindingApiWrapperType) as NetOffice.PowerPointApi.MediaBookmarks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Enums.PpMediaTaskStatus ResamplingStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ResamplingStatus", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpMediaTaskStatus)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public bool IsLinked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsLinked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public bool IsEmbedded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsEmbedded", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 AudioSamplingRate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AudioSamplingRate", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 VideoFrameRate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VideoFrameRate", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 SampleHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SampleHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Int32 SampleWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SampleWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public string VideoCompressionType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VideoCompressionType", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public string AudioCompressionType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AudioCompressionType", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="position">Int32 Position</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void SetDisplayPicture(Int32 position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(position);
			Invoker.Method(this, "SetDisplayPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="filePath">string FilePath</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void SetDisplayPictureFromFile(string filePath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filePath);
			Invoker.Method(this, "SetDisplayPictureFromFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		/// <param name="audioSamplingRate">optional Int32 AudioSamplingRate = 48000</param>
		/// <param name="videoBitRate">optional Int32 VideoBitRate = 7000000</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate, object videoBitRate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate, videoBitRate);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim, object sampleHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim, sampleHeight);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim, object sampleHeight, object sampleWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim, sampleHeight, sampleWidth);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim, sampleHeight, sampleWidth, videoFrameRate);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="trim">optional bool Trim = false</param>
		/// <param name="sampleHeight">optional Int32 SampleHeight = 768</param>
		/// <param name="sampleWidth">optional Int32 SampleWidth = 1280</param>
		/// <param name="videoFrameRate">optional Int32 VideoFrameRate = 24</param>
		/// <param name="audioSamplingRate">optional Int32 AudioSamplingRate = 48000</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Resample(object trim, object sampleHeight, object sampleWidth, object videoFrameRate, object audioSamplingRate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(trim, sampleHeight, sampleWidth, videoFrameRate, audioSamplingRate);
			Invoker.Method(this, "Resample", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="profile">optional NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile = 2</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void ResampleFromProfile(object profile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(profile);
			Invoker.Method(this, "ResampleFromProfile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void ResampleFromProfile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResampleFromProfile", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}