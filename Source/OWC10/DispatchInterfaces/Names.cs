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
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("F5B39BAD-1480-11D3-8549-00C04FAC67D7")]
	public interface Names : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.Name>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.ISpreadsheet Application { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Custom Indexer
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.OWC10Api.Name this[object index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Custom Indexer
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="indexLocal">optional object indexLocal</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.OWC10Api.Name this[object index, object indexLocal] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="indexLocal">optional object indexLocal</param>
		/// <param name="refersTo">optional object refersTo</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.Name this[object index, object indexLocal, object refersTo] { get; }

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Add();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Add(object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Add(object name, object refersTo);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Add(object name, object refersTo, object visible);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Add(object name, object refersTo, object visible, object macroType);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal);

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
		void Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1, object varRefersToR1C1Local);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void AddUI();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void AddUI(object varName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void AddUI(object varName, object varRefersTo);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varName">optional object varName</param>
		/// <param name="varRefersTo">optional object varRefersTo</param>
		/// <param name="varVisible">optional object varVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void AddUI(object varName, object varRefersTo, object varVisible);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal);

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
		void AddUI(object varName, object varRefersTo, object varVisible, object varMacroType, object varShortcutKey, object varCategory, object varNameLocal, object varRefersToLocal, object varCategoryLocal, object varRefersToR1C1);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.Name>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.Name> GetEnumerator();

        #endregion
    }
}
