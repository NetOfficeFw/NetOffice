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
	/// DispatchInterface MailMergeFields 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835151.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MailMergeFields : COMObject ,IEnumerable<NetOffice.WordApi.MailMergeField>
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
                    _type = typeof(MailMergeFields);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMergeFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195330.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834563.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840148.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837178.aspx
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
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.MailMergeField this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836028.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField Add(NetOffice.WordApi.Range range, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		/// <param name="prompt">optional object Prompt</param>
		/// <param name="defaultAskText">optional object DefaultAskText</param>
		/// <param name="askOnce">optional object AskOnce</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText, object askOnce)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name, prompt, defaultAskText, askOnce);
			object returnItem = Invoker.MethodReturn(this, "AddAsk", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name);
			object returnItem = Invoker.MethodReturn(this, "AddAsk", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		/// <param name="prompt">optional object Prompt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name, prompt);
			object returnItem = Invoker.MethodReturn(this, "AddAsk", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		/// <param name="prompt">optional object Prompt</param>
		/// <param name="defaultAskText">optional object DefaultAskText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name, prompt, defaultAskText);
			object returnItem = Invoker.MethodReturn(this, "AddAsk", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="prompt">optional object Prompt</param>
		/// <param name="defaultFillInText">optional object DefaultFillInText</param>
		/// <param name="askOnce">optional object AskOnce</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText, object askOnce)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, prompt, defaultFillInText, askOnce);
			object returnItem = Invoker.MethodReturn(this, "AddFillIn", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddFillIn", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="prompt">optional object Prompt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, prompt);
			object returnItem = Invoker.MethodReturn(this, "AddFillIn", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="prompt">optional object Prompt</param>
		/// <param name="defaultFillInText">optional object DefaultFillInText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, prompt, defaultFillInText);
			object returnItem = Invoker.MethodReturn(this, "AddFillIn", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		/// <param name="trueAutoText">optional object TrueAutoText</param>
		/// <param name="trueText">optional object TrueText</param>
		/// <param name="falseAutoText">optional object FalseAutoText</param>
		/// <param name="falseText">optional object FalseText</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText, object falseText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText, falseText);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		/// <param name="trueAutoText">optional object TrueAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo, trueAutoText);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		/// <param name="trueAutoText">optional object TrueAutoText</param>
		/// <param name="trueText">optional object TrueText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo, trueAutoText, trueText);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		/// <param name="trueAutoText">optional object TrueAutoText</param>
		/// <param name="trueText">optional object TrueText</param>
		/// <param name="falseAutoText">optional object FalseAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText);
			object returnItem = Invoker.MethodReturn(this, "AddIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195621.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddMergeRec(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddMergeRec", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839562.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddMergeSeq(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddMergeSeq", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837747.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNext(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddNext", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo);
			object returnItem = Invoker.MethodReturn(this, "AddNextIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison);
			object returnItem = Invoker.MethodReturn(this, "AddNextIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		/// <param name="valueText">optional object ValueText</param>
		/// <param name="valueAutoText">optional object ValueAutoText</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText, object valueAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name, valueText, valueAutoText);
			object returnItem = Invoker.MethodReturn(this, "AddSet", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name);
			object returnItem = Invoker.MethodReturn(this, "AddSet", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="name">string Name</param>
		/// <param name="valueText">optional object ValueText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, name, valueText);
			object returnItem = Invoker.MethodReturn(this, "AddSet", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		/// <param name="compareTo">optional object CompareTo</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison, compareTo);
			object returnItem = Invoker.MethodReturn(this, "AddSkipIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="mergeField">string MergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison Comparison</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, mergeField, comparison);
			object returnItem = Invoker.MethodReturn(this, "AddSkipIf", paramsArray);
			NetOffice.WordApi.MailMergeField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeField;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.MailMergeField> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.WordApi.MailMergeField> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.MailMergeField item in innerEnumerator)
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