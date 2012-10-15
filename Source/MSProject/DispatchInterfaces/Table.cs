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
	/// DispatchInterface Table 
	/// SupportByVersion MSProject, 11,12,14,15
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Table : COMObject
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
                    _type = typeof(Table);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
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
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
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
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Project newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Project.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Project;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public NetOffice.MSProjectApi.Enums.PjItemType TableType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjItemType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public bool ShowInMenu
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowInMenu", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowInMenu", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public NetOffice.MSProjectApi.Enums.PjDateFormat DateFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSProjectApi.Enums.PjDateFormat)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DateFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public Int32 RowHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public bool LockFirstColumn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockFirstColumn", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockFirstColumn", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public bool AdjustHeaderRowHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AdjustHeaderRowHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AdjustHeaderRowHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public NetOffice.MSProjectApi.TableFields TableFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableFields", paramsArray);
				NetOffice.MSProjectApi.TableFields newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.TableFields.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableFields;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14,15)]
		public void Apply()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Apply", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}