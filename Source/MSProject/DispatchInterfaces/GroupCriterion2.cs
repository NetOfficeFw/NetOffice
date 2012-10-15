using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface GroupCriterion2 
	/// SupportByVersion MSProject, 14,15
	///</summary>
	[SupportByVersionAttribute("MSProject", 14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class GroupCriterion2 : COMObject
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
                    _type = typeof(GroupCriterion2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriterion2(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriterion2(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriterion2(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriterion2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriterion2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.MSProjectApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Application.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public string FieldName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FieldName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Group2 Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Group2 newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Group2.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Group2;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public bool Ascending
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Ascending", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Ascending", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public string FontName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public Int32 FontSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontSize", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public bool FontBold
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontBold", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontBold", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public bool FontItalic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontItalic", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontItalic", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public bool FontUnderLine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontUnderLine", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontUnderLine", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Enums.PjColor FontColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontColor", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjColor)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontColor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Enums.PjColor CellColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CellColor", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjColor)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CellColor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pattern", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjBackgroundPattern)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Pattern", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GroupOn", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjGroupOn)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GroupOn", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public object StartAt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StartAt", paramsArray);
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
				Invoker.PropertySet(this, "StartAt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public object GroupInterval
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GroupInterval", paramsArray);
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
				Invoker.PropertySet(this, "GroupInterval", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public bool Assignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assignment", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Assignment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public Int32 FontColorEx
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontColorEx", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FontColorEx", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public Int32 CellColorEx
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CellColorEx", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CellColorEx", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 14, 15
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14,15)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}