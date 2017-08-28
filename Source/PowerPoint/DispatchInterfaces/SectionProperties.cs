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
 	public class SectionProperties : COMObject
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
                    _type = typeof(SectionProperties);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SectionProperties(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SectionProperties(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745766.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744295.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744380.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746414.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Name(Int32 sectionIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "Name", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745975.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Rename(Int32 sectionIndex, string sectionName)
		{
			 Factory.ExecuteMethod(this, "Rename", sectionIndex, sectionName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745367.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 SlidesCount(Int32 sectionIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "SlidesCount", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744059.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 FirstSlide(Int32 sectionIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "FirstSlide", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745926.aspx </remarks>
		/// <param name="slideIndex">Int32 slideIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 AddBeforeSlide(Int32 slideIndex, string sectionName)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddBeforeSlide", slideIndex, sectionName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">optional object sectionName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 AddSection(Int32 sectionIndex, object sectionName)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddSection", sectionIndex, sectionName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 AddSection(Int32 sectionIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddSection", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746717.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="toPos">Int32 toPos</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Move(Int32 sectionIndex, Int32 toPos)
		{
			 Factory.ExecuteMethod(this, "Move", sectionIndex, toPos);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744948.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="deleteSlides">bool deleteSlides</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Delete(Int32 sectionIndex, bool deleteSlides)
		{
			 Factory.ExecuteMethod(this, "Delete", sectionIndex, deleteSlides);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746673.aspx </remarks>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string SectionID(Int32 sectionIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "SectionID", sectionIndex);
		}

		#endregion

		#pragma warning restore
	}
}
