using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface Slicers 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194125.aspx </remarks>
	public class Slicers : COMObject, NetOffice.ExcelApi.Slicers
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
                    _contractType = typeof(NetOffice.ExcelApi.Slicers);
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
                    _type = typeof(Slicers);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Slicers() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840922.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838612.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194757.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838399.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.ExcelApi.Slicer this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicer>(this, "_Default", typeof(NetOffice.ExcelApi.Slicer), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), new object[]{ slicerDestination, level, name, caption, top, left, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), slicerDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), slicerDestination, level);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), slicerDestination, level, name);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), slicerDestination, level, name, caption);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), new object[]{ slicerDestination, level, name, caption, top });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), new object[]{ slicerDestination, level, name, caption, top, left });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840835.aspx </remarks>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", typeof(NetOffice.ExcelApi.Slicer), new object[]{ slicerDestination, level, name, caption, top, left, width });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Slicer>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Slicer>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Slicer>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Slicer>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.Slicer> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Slicer item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

