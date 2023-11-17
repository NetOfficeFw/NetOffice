using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IQueries 
	/// SupportByVersion Excel, 16
	/// </summary>
	[SupportByVersion("Excel", 16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class IQueries : COMObject, IEnumerableProvider<NetOffice.ExcelApi.WorkbookQuery>
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
					_type = typeof(IQueries);
				return _type;
			}
		}
		
		#endregion
		
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IQueries(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		public IQueries(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IQueries(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.WorkbookQuery this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookQuery>(this, "_Default", NetOffice.ExcelApi.WorkbookQuery.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="name">name of the query</param>
		/// <param name="formula">Power Query M formula for the new query</param>
		/// <param name="description">optional description of the query</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.WorkbookQuery Add(string name, string formula, object description)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookQuery>(this, "Add", NetOffice.ExcelApi.WorkbookQuery.LateBindingApiWrapperType, name, formula, description);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="name">name of the query</param>
		/// <param name="formula">Power Query M formula for the new query</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.WorkbookQuery Add(string name, string formula)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookQuery>(this, "Add", NetOffice.ExcelApi.WorkbookQuery.LateBindingApiWrapperType, name, formula);
		}

		#endregion

		#region IEnumerableProvider<NetOffice.ExcelApi.WorkbookQuery>

		ICOMObject IEnumerableProvider<NetOffice.ExcelApi.WorkbookQuery>.GetComObjectEnumerator(ICOMObject parent)
		{
			return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
		}

		IEnumerable IEnumerableProvider<NetOffice.ExcelApi.WorkbookQuery>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
		{
			return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
		}

		#endregion

		#region IEnumerable<NetOffice.ExcelApi.WorkbookQuery>

		/// <summary>
		/// SupportByVersion Excel, 16
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public IEnumerator<NetOffice.ExcelApi.WorkbookQuery> GetEnumerator()
		{
			NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
			foreach (NetOffice.ExcelApi.WorkbookQuery item in innerEnumerator)
				yield return item;
		}

		#endregion

		#region IEnumerable

		/// <summary>
		/// SupportByVersion Excel, 16
		/// </summary>
		[SupportByVersion("Excel", 16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}