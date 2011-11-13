using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.ADODBApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
	[ComImport, Guid("00000266-0000-0010-8000-00AA006D2EA4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RecordsetEvents
	{
		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void WillChangeField([In] object cFields, [In, MarshalAs(UnmanagedType.IDispatch)] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void FieldChangeComplete([In] object cFields, [In, MarshalAs(UnmanagedType.IDispatch)] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByLibraryAttribute("ADODB", 2.1,2.5)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class RecordsetEvents_SinkHelper : SinkHelper, RecordsetEvents
	{
		#region Static
		
		public static readonly string Id = "00000266-0000-0010-8000-00AA006D2EA4";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public RecordsetEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region RecordsetEvents Members
		
		public void WillChangeField([In] object cFields, [In, MarshalAs(UnmanagedType.IDispatch)] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WillChangeField");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cFields, fields, adStatus, pRecordset);
				return;
			}

			Int32 newcFields = (Int32)cFields;
			object newFields = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, fields) as object;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[4];
			paramsArray[0] = newcFields;
			paramsArray[1] = newFields;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void FieldChangeComplete([In] object cFields, [In, MarshalAs(UnmanagedType.IDispatch)] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FieldChangeComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cFields, fields, pError, adStatus, pRecordset);
				return;
			}

			Int32 newcFields = (Int32)cFields;
			object newFields = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, fields) as object;
			NetOffice.ADODBApi.Error newpError = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pError) as NetOffice.ADODBApi.Error;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[5];
			paramsArray[0] = newcFields;
			paramsArray[1] = newFields;
			paramsArray[2] = newpError;
			paramsArray[3] = newadStatus;
			paramsArray[4] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WillChangeRecord");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, cRecords, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			Int32 newcRecords = (Int32)cRecords;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newcRecords;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RecordChangeComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, cRecords, pError, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			Int32 newcRecords = (Int32)cRecords;
			NetOffice.ADODBApi.Error newpError = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pError) as NetOffice.ADODBApi.Error;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[5];
			paramsArray[0] = newadReason;
			paramsArray[1] = newcRecords;
			paramsArray[2] = newpError;
			paramsArray[3] = newadStatus;
			paramsArray[4] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WillChangeRecordset");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[3];
			paramsArray[0] = newadReason;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RecordsetChangeComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Error newpError = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pError) as NetOffice.ADODBApi.Error;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WillMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[3];
			paramsArray[0] = newadReason;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MoveComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Error newpError = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pError) as NetOffice.ADODBApi.Error;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EndOfRecordset");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(fMoreData, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[3];
			paramsArray.SetValue(fMoreData, 0);
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			fMoreData = (bool)paramsArray[0];
		}

		public void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FetchProgress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(progress, maxProgress, adStatus, pRecordset);
				return;
			}

			Int32 newProgress = (Int32)progress;
			Int32 newMaxProgress = (Int32)maxProgress;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[4];
			paramsArray[0] = newProgress;
			paramsArray[1] = newMaxProgress;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FetchComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pError, adStatus, pRecordset);
				return;
			}

			NetOffice.ADODBApi.Error newpError = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pError) as NetOffice.ADODBApi.Error;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
			NetOffice.ADODBApi._Recordset newpRecordset = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
			object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}