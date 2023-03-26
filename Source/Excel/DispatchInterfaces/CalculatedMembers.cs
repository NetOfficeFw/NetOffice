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
	/// DispatchInterface CalculatedMembers 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class CalculatedMembers : COMObject, IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>
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
                    _type = typeof(CalculatedMembers);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CalculatedMembers(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CalculatedMembers(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CalculatedMembers(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Application"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Creator"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Parent"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Count"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.CalculatedMember this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Default", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder, type);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="dynamic">optional object dynamic</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		/// <param name="hierarchizeDistinct">optional object hierarchizeDistinct</param>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder, object hierarchizeDistinct)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, dynamic, displayFolder, hierarchizeDistinct });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, string formula)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="dynamic">optional object dynamic</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, dynamic });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CalculatedMembers.Add"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="dynamic">optional object dynamic</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, dynamic, displayFolder });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder, type);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="formula">string formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		/// <param name="measureGroup">optional object measureGroup</param>
		/// <param name="parentHierarchy">optional object parentHierarchy</param>
		/// <param name="parentMember">optional object parentMember</param>
		/// <param name="numberFormat">optional object numberFormat</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember, object numberFormat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, parentMember, numberFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, name, formula, solveOrder, type);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, displayFolder });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		/// <param name="measureGroup">optional object measureGroup</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, displayFolder, measureGroup });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		/// <param name="measureGroup">optional object measureGroup</param>
		/// <param name="parentHierarchy">optional object parentHierarchy</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.calculatedmembers.addcalculatedmember"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="formula">object formula</param>
		/// <param name="solveOrder">optional object solveOrder</param>
		/// <param name="type">optional object type</param>
		/// <param name="displayFolder">optional object displayFolder</param>
		/// <param name="measureGroup">optional object measureGroup</param>
		/// <param name="parentHierarchy">optional object parentHierarchy</param>
		/// <param name="parentMember">optional object parentMember</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", NetOffice.ExcelApi.CalculatedMember.LateBindingApiWrapperType, new object[]{ name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, parentMember });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.CalculatedMember>

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.CalculatedMember> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.CalculatedMember item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}