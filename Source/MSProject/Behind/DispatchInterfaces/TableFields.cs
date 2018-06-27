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
	/// DispatchInterface TableFields 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920711(v=office.14).aspx </remarks>
	public class TableFields : COMObject, NetOffice.MSProjectApi.TableFields
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
                    _contractType = typeof(NetOffice.MSProjectApi.TableFields);
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
                    _type = typeof(TableFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TableFields() : base()
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
		public virtual NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Project>(this, "Parent", typeof(NetOffice.MSProjectApi.Project));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.MSProjectApi.TableField this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TableField>(this, "Item", typeof(NetOffice.MSProjectApi.TableField), index);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		/// <param name="autoWrap">optional bool AutoWrap = true</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before, object autoWrap)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), new object[]{ field, alignData, width, title, alignTitle, before, autoWrap });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), field);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), field, alignData);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), field, alignData, width);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), field, alignData, width, title);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), new object[]{ field, alignData, width, title, alignTitle });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", typeof(NetOffice.MSProjectApi.TableField), new object[]{ field, alignData, width, title, alignTitle, before });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.TableField>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.TableField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.TableField>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.TableField>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        public virtual IEnumerator<NetOffice.MSProjectApi.TableField> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.TableField item in innerEnumerator)
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

