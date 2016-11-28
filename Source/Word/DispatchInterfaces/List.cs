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
	/// DispatchInterface List 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192789.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class List : COMObject
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
                    _type = typeof(List);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public List(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840272.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193376.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListParagraphs", paramsArray);
				NetOffice.WordApi.ListParagraphs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType) as NetOffice.WordApi.ListParagraphs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838053.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194676.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836145.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194035.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194057.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string StyleName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx
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
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196826.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, defaultListBehavior);
			Invoker.Method(this, "ApplyListTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		/// <param name="applyLevel">optional object ApplyLevel</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior, object applyLevel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, defaultListBehavior, applyLevel);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate ListTemplate</param>
		/// <param name="continuePreviousList">optional object ContinuePreviousList</param>
		/// <param name="defaultListBehavior">optional object DefaultListBehavior</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listTemplate, continuePreviousList, defaultListBehavior);
			Invoker.Method(this, "ApplyListTemplateWithLevel", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}