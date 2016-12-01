using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface Names 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Names : COMObject ,IEnumerable<NetOffice.OWC10Api.Name>
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
                    _type = typeof(Names);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Names(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OWC10Api.ISpreadsheet newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api.ISpreadsheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="indexLocal">optional object IndexLocal</param>
		/// <param name="refersTo">optional object RefersTo</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OWC10Api.Name this[object index, object indexLocal, object refersTo]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index, indexLocal, refersTo);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.OWC10Api.Name newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.Name.LateBindingApiWrapperType) as NetOffice.OWC10Api.Name;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="nameLocal">optional object NameLocal</param>
		/// <param name="refersToLocal">optional object RefersToLocal</param>
		/// <param name="categoryLocal">optional object CategoryLocal</param>
		/// <param name="refersToR1C1">optional object RefersToR1C1</param>
		/// <param name="refersToR1C1Local">optional object RefersToR1C1Local</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="nameLocal">optional object NameLocal</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category, nameLocal);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="nameLocal">optional object NameLocal</param>
		/// <param name="refersToLocal">optional object RefersToLocal</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="nameLocal">optional object NameLocal</param>
		/// <param name="refersToLocal">optional object RefersToLocal</param>
		/// <param name="categoryLocal">optional object CategoryLocal</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="refersTo">optional object RefersTo</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="macroType">optional object MacroType</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="nameLocal">optional object NameLocal</param>
		/// <param name="refersToLocal">optional object RefersToLocal</param>
		/// <param name="categoryLocal">optional object CategoryLocal</param>
		/// <param name="refersToR1C1">optional object RefersToR1C1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1);
			Invoker.Method(this, "Add", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		/// <param name="varRefersToLocal">optional object varRefersToLocal</param>
		/// <param name="varCategoryLocal">optional object varCategoryLocal</param>
		/// <param name="varRefersToR1C1">optional object varRefersToR1C1</param>
		/// <param name="varRefersToR1C1Local">optional object varRefersToR1C1Local</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1, object varRefersToR1C1Local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1, varRefersToR1C1Local);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		/// <param name="varRefersToLocal">optional object varRefersToLocal</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		/// <param name="varRefersToLocal">optional object varRefersToLocal</param>
		/// <param name="varCategoryLocal">optional object varCategoryLocal</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		/// <param name="varRefersToLocal">optional object varRefersToLocal</param>
		/// <param name="varCategoryLocal">optional object varCategoryLocal</param>
		/// <param name="varRefersToR1C1">optional object varRefersToR1C1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1);
			Invoker.Method(this, "AddUI", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OWC10Api.Name> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.Name> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.Name item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}