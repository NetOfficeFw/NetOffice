using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface AnimationSettings 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746630.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AnimationSettings : COMObject
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
                    _type = typeof(AnimationSettings);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AnimationSettings(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AnimationSettings(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AnimationSettings(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746265.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745262.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744926.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorFormat DimColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorFormat>(this, "DimColor", NetOffice.PowerPointApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745985.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SoundEffect SoundEffect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SoundEffect>(this, "SoundEffect", NetOffice.PowerPointApi.SoundEffect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746472.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpEntryEffect EntryEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpEntryEffect>(this, "EntryEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EntryEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746411.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpAfterEffect AfterEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpAfterEffect>(this, "AfterEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AfterEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743950.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 AnimationOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AnimationOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AnimationOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745337.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpAdvanceMode AdvanceMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpAdvanceMode>(this, "AdvanceMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AdvanceMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746720.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single AdvanceTime
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "AdvanceTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AdvanceTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744348.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PlaySettings PlaySettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PlaySettings>(this, "PlaySettings", NetOffice.PowerPointApi.PlaySettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743834.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpTextLevelEffect TextLevelEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpTextLevelEffect>(this, "TextLevelEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextLevelEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745149.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpTextUnitEffect TextUnitEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpTextUnitEffect>(this, "TextUnitEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextUnitEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745264.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Animate
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Animate");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Animate", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745644.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AnimateBackground
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AnimateBackground");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AnimateBackground", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746284.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AnimateTextInReverse
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AnimateTextInReverse");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AnimateTextInReverse", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745825.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpChartUnitEffect ChartUnitEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpChartUnitEffect>(this, "ChartUnitEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ChartUnitEffect", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
