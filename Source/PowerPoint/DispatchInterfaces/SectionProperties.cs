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
	/// DispatchInterface SectionProperties 
	/// SupportByVersion PowerPoint, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743911.aspx
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SectionProperties : COMObject
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
                    _type = typeof(SectionProperties);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SectionProperties(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745766.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744295.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744380.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746414.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public string Name(Int32 sectionIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex);
			object returnItem = Invoker.MethodReturn(this, "Name", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745975.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void Rename(Int32 sectionIndex, string sectionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex, sectionName);
			Invoker.Method(this, "Rename", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745367.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 SlidesCount(Int32 sectionIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex);
			object returnItem = Invoker.MethodReturn(this, "SlidesCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744059.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 FirstSlide(Int32 sectionIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex);
			object returnItem = Invoker.MethodReturn(this, "FirstSlide", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745926.aspx
		/// </summary>
		/// <param name="slideIndex">Int32 SlideIndex</param>
		/// <param name="sectionName">string sectionName</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 AddBeforeSlide(Int32 slideIndex, string sectionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideIndex, sectionName);
			object returnItem = Invoker.MethodReturn(this, "AddBeforeSlide", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="sectionName">optional object sectionName</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 AddSection(Int32 sectionIndex, object sectionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex, sectionName);
			object returnItem = Invoker.MethodReturn(this, "AddSection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746122.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public Int32 AddSection(Int32 sectionIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex);
			object returnItem = Invoker.MethodReturn(this, "AddSection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746717.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="toPos">Int32 toPos</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void Move(Int32 sectionIndex, Int32 toPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex, toPos);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744948.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="deleteSlides">bool deleteSlides</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void Delete(Int32 sectionIndex, bool deleteSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex, deleteSlides);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746673.aspx
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public string SectionID(Int32 sectionIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectionIndex);
			object returnItem = Invoker.MethodReturn(this, "SectionID", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}