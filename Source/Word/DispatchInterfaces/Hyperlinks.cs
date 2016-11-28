using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Hyperlinks 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192421.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Hyperlinks : COMObject ,IEnumerable<NetOffice.WordApi.Hyperlink>
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
                    _type = typeof(Hyperlinks);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838934.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197858.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837725.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195476.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.Hyperlink this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor, object address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address, subAddress);
			object returnItem = Invoker.MethodReturn(this, "_Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor);
			object returnItem = Invoker.MethodReturn(this, "_Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address);
			object returnItem = Invoker.MethodReturn(this, "_Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="screenTip">optional object ScreenTip</param>
		/// <param name="textToDisplay">optional object TextToDisplay</param>
		/// <param name="target">optional object Target</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip, object textToDisplay, object target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address, subAddress, screenTip, textToDisplay, target);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address, subAddress);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="screenTip">optional object ScreenTip</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address, subAddress, screenTip);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837214.aspx
		/// </summary>
		/// <param name="anchor">object Anchor</param>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="screenTip">optional object ScreenTip</param>
		/// <param name="textToDisplay">optional object TextToDisplay</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip, object textToDisplay)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, address, subAddress, screenTip, textToDisplay);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlink;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.Hyperlink> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.WordApi.Hyperlink> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.Hyperlink item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}