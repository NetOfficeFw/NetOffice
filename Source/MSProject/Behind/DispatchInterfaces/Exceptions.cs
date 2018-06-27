using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface Exceptions 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920590(v=office.14).aspx </remarks>
	public class Exceptions : COMObject, NetOffice.MSProjectApi.Exceptions
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
                    _contractType = typeof(NetOffice.MSProjectApi.Exceptions);
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
                    _type = typeof(Exceptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Exceptions() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", typeof(NetOffice.MSProjectApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Calendar Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "Parent", typeof(NetOffice.MSProjectApi.Calendar));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.MSProjectApi.Exception this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Exception>(this, "Item", typeof(NetOffice.MSProjectApi.Exception), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		/// <param name="monthDay">optional object monthDay</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month, object monthDay)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month, monthDay });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), type, start);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), type, start, finish);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), type, start, finish, occurrences);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period, daysOfWeek });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", typeof(NetOffice.MSProjectApi.Exception), new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.Exception>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.Exception>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.Exception>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Exception>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        public virtual IEnumerator<NetOffice.MSProjectApi.Exception> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.Exception item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

