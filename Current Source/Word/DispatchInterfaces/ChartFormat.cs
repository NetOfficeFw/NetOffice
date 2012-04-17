using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface ChartFormat 
	/// SupportByVersion Word, 14
	///</summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ChartFormat : COMObject
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
                    _type = typeof(ChartFormat);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartFormat(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartFormat(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartFormat() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.FillFormat Fill
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fill", paramsArray);
				NetOffice.WordApi.FillFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FillFormat.LateBindingApiWrapperType) as NetOffice.WordApi.FillFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.GlowFormat Glow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Glow", paramsArray);
				NetOffice.WordApi.GlowFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.GlowFormat.LateBindingApiWrapperType) as NetOffice.WordApi.GlowFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.LineFormat Line
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Line", paramsArray);
				NetOffice.WordApi.LineFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.LineFormat.LateBindingApiWrapperType) as NetOffice.WordApi.LineFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.PictureFormat PictureFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PictureFormat", paramsArray);
				NetOffice.WordApi.PictureFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.PictureFormat.LateBindingApiWrapperType) as NetOffice.WordApi.PictureFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.ShadowFormat Shadow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shadow", paramsArray);
				NetOffice.WordApi.ShadowFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ShadowFormat.LateBindingApiWrapperType) as NetOffice.WordApi.ShadowFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.SoftEdgeFormat SoftEdge
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SoftEdge", paramsArray);
				NetOffice.WordApi.SoftEdgeFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SoftEdgeFormat.LateBindingApiWrapperType) as NetOffice.WordApi.SoftEdgeFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.OfficeApi.TextFrame2 TextFrame2
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextFrame2", paramsArray);
				NetOffice.OfficeApi.TextFrame2 newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.TextFrame2.LateBindingApiWrapperType) as NetOffice.OfficeApi.TextFrame2;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.ThreeDFormat ThreeD
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThreeD", paramsArray);
				NetOffice.WordApi.ThreeDFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ThreeDFormat.LateBindingApiWrapperType) as NetOffice.WordApi.ThreeDFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public object Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		#endregion
		#pragma warning restore
	}
}