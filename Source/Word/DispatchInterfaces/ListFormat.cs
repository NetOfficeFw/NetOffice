using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface ListFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820923.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ListFormat : COMObject
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
                    _type = typeof(ListFormat);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ListFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844779.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ListLevelNumber
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListLevelNumber", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ListLevelNumber", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839514.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.List List
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "List", paramsArray);
				NetOffice.WordApi.List newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.List.LateBindingApiWrapperType) as NetOffice.WordApi.List;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821156.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListTemplate ListTemplate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListTemplate", paramsArray);
				NetOffice.WordApi.ListTemplate newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListTemplate.LateBindingApiWrapperType) as NetOffice.WordApi.ListTemplate;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196311.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ListValue
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListValue", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836742.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SingleList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SingleList", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835216.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SingleListTemplate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SingleListTemplate", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197765.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdListType ListType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdListType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836909.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ListString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194491.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195720.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194230.aspx
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837259.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape ListPictureBullet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListPictureBullet", paramsArray);
				NetOffice.WordApi.InlineShape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196639.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdContinue CanContinuePreviousList(NetOffice.WordApi.ListTemplate listTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate);
			object returnItem = Invoker.MethodReturn(this, "CanContinuePreviousList", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.WordApi.Enums.WdContinue)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			Invoker.Method(this, "RemoveNumbers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveNumbers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			Invoker.Method(this, "ConvertNumbersToText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertNumbersToText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		/// <param name="level">optional object Level</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType, object level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType, level);
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefaultOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyBulletDefaultOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefaultOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyNumberDefaultOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefaultOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyOutlineNumberDefaultOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo);
			Invoker.Method(this, "ApplyListTemplateOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate);
			Invoker.Method(this, "ApplyListTemplateOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList);
			Invoker.Method(this, "ApplyListTemplateOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840622.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ListOutdent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ListOutdent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192794.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ListIndent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ListIndent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx
		/// </summary>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefault(object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(defaultListBehavior);
			Invoker.Method(this, "ApplyBulletDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefault()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyBulletDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx
		/// </summary>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefault(object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(defaultListBehavior);
			Invoker.Method(this, "ApplyNumberDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefault()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyNumberDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx
		/// </summary>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefault(object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(defaultListBehavior);
			Invoker.Method(this, "ApplyOutlineNumberDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefault()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyOutlineNumberDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo, defaultListBehavior);
			Invoker.Method(this, "ApplyListTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate);
			Invoker.Method(this, "ApplyListTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList);
			Invoker.Method(this, "ApplyListTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo);
			Invoker.Method(this, "ApplyListTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		/// <param name="applyLevel">optional object ApplyLevel</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior, object applyLevel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo, defaultListBehavior, applyLevel);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="applyTo">optional object ApplyTo</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, applyTo, defaultListBehavior);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}