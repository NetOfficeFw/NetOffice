using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface ResampleMediaTask 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746356.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ResampleMediaTask : COMObject
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
                    _type = typeof(ResampleMediaTask);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ResampleMediaTask(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ResampleMediaTask(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ResampleMediaTask(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745595.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpResampleMediaProfile>(this, "profile");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745655.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744051.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745190.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744753.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743901.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745487.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746165.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string ContainerType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ContainerType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746039.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744272.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string AudioCompressionType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AudioCompressionType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744125.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "Shape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
