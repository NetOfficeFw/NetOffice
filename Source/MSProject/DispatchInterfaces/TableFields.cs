using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface TableFields 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920711(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class TableFields : COMObject ,IEnumerable<NetOffice.MSProjectApi.TableField>
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
                    _type = typeof(TableFields);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TableFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.MSProjectApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Application.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Project newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Project.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Project;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSProjectApi.TableField this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		/// <param name="autoWrap">optional bool AutoWrap = true</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before, object autoWrap)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData, width, title, alignTitle, before, autoWrap);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData, width);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData, width, title);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData, width, title, alignTitle);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, alignData, width, title, alignTitle, before);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.TableField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType) as NetOffice.MSProjectApi.TableField;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.TableField> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
       public IEnumerator<NetOffice.MSProjectApi.TableField> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.TableField item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}