using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface Names 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Names : COMObject, IEnumerableProvider<NetOffice.OWC10Api.Name>
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Names(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

        #endregion

        #region Methods
        
        /// <summary>
        /// SupportByVersion OWC10 1
        /// Custom Indexer
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public NetOffice.OWC10Api.Name this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", NetOffice.OWC10Api.Name.LateBindingApiWrapperType, index);
			}
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Custom Indexer
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="indexLocal">optional object indexLocal</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public NetOffice.OWC10Api.Name this[object index, object indexLocal]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", NetOffice.OWC10Api.Name.LateBindingApiWrapperType, index, indexLocal);
			}
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="indexLocal">optional object indexLocal</param>
        /// <param name="refersTo">optional object refersTo</param>
        [SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OWC10Api.Name this[object index, object indexLocal, object refersTo]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", NetOffice.OWC10Api.Name.LateBindingApiWrapperType, index, indexLocal, refersTo);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		/// <param name="refersToR1C1Local">optional object refersToR1C1Local</param>
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add()
		{
			 Factory.ExecuteMethod(this, "Add");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name)
		{
			 Factory.ExecuteMethod(this, "Add", name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo)
		{
			 Factory.ExecuteMethod(this, "Add", name, refersTo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible)
		{
			 Factory.ExecuteMethod(this, "Add", name, refersTo, visible);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType)
		{
			 Factory.ExecuteMethod(this, "Add", name, refersTo, visible, macroType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1)
		{
			 Factory.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1 });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1, object varRefersToR1C1Local)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1, varRefersToR1C1Local });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI()
		{
			 Factory.ExecuteMethod(this, "AddUI");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName)
		{
			 Factory.ExecuteMethod(this, "AddUI", varName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo)
		{
			 Factory.ExecuteMethod(this, "AddUI", varName, varRefersTo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible)
		{
			 Factory.ExecuteMethod(this, "AddUI", varName, varRefersTo, varVisible);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType)
		{
			 Factory.ExecuteMethod(this, "AddUI", varName, varRefersTo, varVisible, varMacroType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		/// <param name="varMacroType">optional object varMacroType</param>
		/// <param name="varShortcutKey">optional object varShortcutKey</param>
		/// <param name="varCategory">optional object varCategory</param>
		/// <param name="varNameLocal">optional object varNameLocal</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1)
		{
			 Factory.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1 });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.Name>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.Name>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.Name>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.Name>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<NetOffice.OWC10Api.Name> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.Name item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}