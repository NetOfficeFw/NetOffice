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
	/// DispatchInterface Footnotes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838736.aspx </remarks>
	public class Footnotes : COMObject, NetOffice.WordApi.Footnotes
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
                    _contractType = typeof(NetOffice.WordApi.Footnotes);
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
                    _type = typeof(Footnotes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Footnotes() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192001.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834874.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192141.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837505.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdFootnoteLocation Location
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFootnoteLocation>(this, "Location");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Location", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840012.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdNoteNumberStyle NumberStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNoteNumberStyle>(this, "NumberStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "NumberStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837690.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 StartingNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StartingNumber");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartingNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838274.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdNumberingRule NumberingRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNumberingRule>(this, "NumberingRule");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "NumberingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821261.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Separator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Separator", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196589.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ContinuationSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "ContinuationSeparator", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193419.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ContinuationNotice
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "ContinuationNotice", typeof(NetOffice.WordApi.Range));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.WordApi.Footnote this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Footnote>(this, "Item", typeof(NetOffice.WordApi.Footnote), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834548.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="reference">optional object reference</param>
		/// <param name="text">optional object text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Footnote Add(NetOffice.WordApi.Range range, object reference, object text)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Footnote>(this, "Add", typeof(NetOffice.WordApi.Footnote), range, reference, text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834548.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Footnote Add(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Footnote>(this, "Add", typeof(NetOffice.WordApi.Footnote), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834548.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="reference">optional object reference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Footnote Add(NetOffice.WordApi.Range range, object reference)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Footnote>(this, "Add", typeof(NetOffice.WordApi.Footnote), range, reference);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198007.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Convert()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Convert");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838265.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SwapWithEndnotes()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SwapWithEndnotes");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192404.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ResetSeparator()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResetSeparator");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840119.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ResetContinuationSeparator()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResetContinuationSeparator");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821339.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ResetContinuationNotice()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResetContinuationNotice");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Footnote>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Footnote>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Footnote>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Footnote>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.Footnote> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Footnote item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

