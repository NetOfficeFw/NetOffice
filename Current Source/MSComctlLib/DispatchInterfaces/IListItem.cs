using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSComctlLibApi
{
	///<summary>
	/// DispatchInterface IListItem 
	/// SupportByVersion MSComctlLib, 6.0
	///</summary>
	[SupportByVersionAttribute("MSComctlLib", 6.0)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IListItem : COMObject
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
                    _type = typeof(IListItem);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItem(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItem(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItem(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItem() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public string Default
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Default", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Default", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public string Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Text", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public bool Ghosted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Ghosted", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Ghosted", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Single Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Height", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public object Icon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Icon", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Icon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Index", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public string Key
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Key", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Key", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Single Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Left", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public bool Selected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selected", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Selected", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public object SmallIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmallIcon", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SmallIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public object Tag
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tag", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Tag", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Single Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Top", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Single Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Width", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_SubItems(Int16 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "SubItems", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public void set_SubItems(Int16 index, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.PropertySet(this, "SubItems", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Alias for get_SubItems
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public string SubItems(Int16 index)
		{
			return get_SubItems(index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public NetOffice.MSComctlLibApi.IListSubItems ListSubItems
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListSubItems", paramsArray);
				NetOffice.MSComctlLibApi.IListSubItems newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IListSubItems;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ListSubItems", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public bool Checked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Checked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Checked", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public Int32 ForeColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ForeColor", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ForeColor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public string ToolTipText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ToolTipText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ToolTipText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public bool Bold
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bold", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Bold", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public stdole.Picture CreateDragImage()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateDragImage", paramsArray);
			stdole.Picture newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem) as stdole.Picture;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6.0)]
		public bool EnsureVisible()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EnsureVisible", paramsArray);
			return (bool)returnItem;
		}

		#endregion
		#pragma warning restore
	}
}