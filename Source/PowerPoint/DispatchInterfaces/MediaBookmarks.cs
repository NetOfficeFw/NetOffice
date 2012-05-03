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
	/// DispatchInterface MediaBookmarks 
	/// SupportByVersion PowerPoint, 14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MediaBookmarks : Collection
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
                    _type = typeof(MediaBookmarks);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaBookmarks(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaBookmarks(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaBookmarks(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaBookmarks() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MediaBookmarks(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PowerPointApi.MediaBookmark this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PowerPointApi.MediaBookmark newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.MediaBookmark.LateBindingApiWrapperType) as NetOffice.PowerPointApi.MediaBookmark;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="position">Int32 Position</param>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.MediaBookmark Add(Int32 position, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(position, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.MediaBookmark newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.MediaBookmark.LateBindingApiWrapperType) as NetOffice.PowerPointApi.MediaBookmark;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}