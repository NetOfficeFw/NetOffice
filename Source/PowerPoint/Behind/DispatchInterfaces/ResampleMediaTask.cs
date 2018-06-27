using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface ResampleMediaTask 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746356.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ResampleMediaTask : COMObject, NetOffice.PowerPointApi.ResampleMediaTask
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
                    _contractType = typeof(NetOffice.PowerPointApi.ResampleMediaTask);
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
                    _type = typeof(ResampleMediaTask);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ResampleMediaTask() : base()
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
		public NetOffice.PowerPointApi.Enums.PpResampleMediaProfile profile		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpResampleMediaProfile>(this, "profile");
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
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsLinked");
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
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsEmbedded");
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AudioSamplingRate");
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "VideoFrameRate");
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SampleHeight");
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SampleWidth");
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
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ContainerType");
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
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "VideoCompressionType");
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
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AudioCompressionType");
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "Shape", typeof(NetOffice.PowerPointApi.Shape));
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


