using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface OMathFunctions 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845910.aspx </remarks>
	public class OMathFunctions : COMObject, NetOffice.WordApi.OMathFunctions
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
                    _contractType = typeof(NetOffice.WordApi.OMathFunctions);
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
                    _type = typeof(OMathFunctions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public OMathFunctions() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823229.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838138.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837013.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836429.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.WordApi.OMathFunction this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.OMathFunction>(this, "Item", typeof(NetOffice.WordApi.OMathFunction), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		/// <param name="numArgs">optional object numArgs</param>
		/// <param name="numCols">optional object numCols</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type, object numArgs, object numCols)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.OMathFunction>(this, "Add", typeof(NetOffice.WordApi.OMathFunction), range, type, numArgs, numCols);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.OMathFunction>(this, "Add", typeof(NetOffice.WordApi.OMathFunction), range, type);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		/// <param name="numArgs">optional object numArgs</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type, object numArgs)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.OMathFunction>(this, "Add", typeof(NetOffice.WordApi.OMathFunction), range, type, numArgs);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.OMathFunction>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.OMathFunction>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.OMathFunction>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.OMathFunction>

        /// <summary>
        /// SupportByVersion Word, 12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.OMathFunction> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.OMathFunction item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

