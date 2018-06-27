using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface Names 
	/// SupportByVersion OWC10, 1
	/// </summary>
	public class Names : COMObject, NetOffice.OWC10Api.Names
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OWC10Api.Names);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Names() : base()
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
		public virtual NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
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
        public virtual NetOffice.OWC10Api.Name this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", typeof(NetOffice.OWC10Api.Name), index);
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
        public virtual NetOffice.OWC10Api.Name this[object index, object indexLocal]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", typeof(NetOffice.OWC10Api.Name), index, indexLocal);
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
		public virtual NetOffice.OWC10Api.Name this[object index, object indexLocal, object refersTo]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Name>(this, "Item", typeof(NetOffice.OWC10Api.Name), index, indexLocal, refersTo);
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Add()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Add(object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Add(object name, object refersTo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", name, refersTo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Add(object name, object refersTo, object visible)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", name, refersTo, visible);
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
		public virtual void Add(object name, object refersTo, object visible, object macroType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", name, refersTo, visible, macroType);
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey });
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category });
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal });
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal });
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal });
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
		public virtual void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1 });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1, object varRefersToR1C1Local)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1, varRefersToR1C1Local });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void AddUI()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void AddUI(object varName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", varName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void AddUI(object varName, object varRefersTo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", varName, varRefersTo);
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", varName, varRefersTo, varVisible);
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", varName, varRefersTo, varVisible, varMacroType);
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal });
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
		public virtual void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddUI", new object[]{ varName, varRefersTo, varVisible, varMacroType, varShortcutKey, varCategory, varNameLocal, varRefersToLocal, varCategoryLocal, varRefersToR1C1 });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.Name>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.Name>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.OWC10Api.Name> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

