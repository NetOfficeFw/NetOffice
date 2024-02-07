using System;
using System.Collections.Generic;
using NetOffice.Availability;

namespace NetOffice.Tests.Helpers
{
    internal class TestableComObjectStub : ICOMObject, IEventBinding
    {
        private List<ICOMObject> _children = new List<ICOMObject>();
        private List<string> _eventRecipients = new List<string>();

        public void AddEventRecipient(string eventName)
        {
            this._eventRecipients.Add(eventName);
        }

        public bool EventBridgeInitialized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public Core Factory => Core.Default;

        public Invoker Invoker => throw new NotImplementedException();

        public Settings Settings => Core.Default.Settings;

        public DebugConsole Console => throw new NotImplementedException();

        public object UnderlyingObject => throw new NotImplementedException();

        public Type UnderlyingType => throw new NotImplementedException();

        public string UnderlyingTypeName => throw new NotImplementedException();

        public string UnderlyingFriendlyTypeName => throw new NotImplementedException();

        public string UnderlyingComponentName => throw new NotImplementedException();

        public string InstanceName => throw new NotImplementedException();

        public string InstanceFriendlyName => throw new NotImplementedException();

        public string InstanceComponentName => throw new NotImplementedException();

        public Type InstanceType => throw new NotImplementedException();

        public bool IsDisposed => false;

        public bool IsCurrentlyDisposing => false;

        public ICOMObject ParentObject => throw new NotImplementedException();

        public IEnumerable<ICOMObject> ChildObjects => _children;

        public bool IsEventBinding => true;

        public bool IsEventBridgeInitialized => true;

        public bool IsWithEventRecipients => throw new NotImplementedException();

        public event OnDisposeEventHandler OnDispose;

        public void AddChildObject(ICOMObject childObject)
        {
            _children.Add(childObject);
        }

        public object Clone()
        {
            throw new NotImplementedException();
        }

        public void CreateEventBridge()
        {
            throw new NotImplementedException();
        }

        public void Dispose(bool disposeEventBinding)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public void DisposeChildInstances()
        {
            throw new NotImplementedException();
        }

        public void DisposeChildInstances(bool disposeEventBinding)
        {
            throw new NotImplementedException();
        }

        public void DisposeEventBridge()
        {
            throw new NotImplementedException();
        }

        public bool EntityIsAvailable(string name)
        {
            throw new NotImplementedException();
        }

        public bool EntityIsAvailable(string name, SupportedEntityType searchType)
        {
            throw new NotImplementedException();
        }

        public int GetCountOfEventRecipients(string eventName)
        {
            throw new NotImplementedException();
        }

        public Delegate[] GetEventRecipients(string eventName)
        {
            throw new NotImplementedException();
        }

        public bool HasEventRecipients()
        {
            return this._eventRecipients.Count > 0;
        }

        public bool HasEventRecipients(string eventName)
        {
            return this._eventRecipients.Contains(eventName);
        }

        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
        {
            this.LastRaisedEventName = eventName;
            this.LastRaisedEventParameters = paramsArray;

            return 0;
        }

        public bool RemoveChildObject(ICOMObject childObject)
        {
            return _children.Remove(childObject);
        }

        T ICOMObject.To<T>()
        {
            throw new NotImplementedException();
        }

        public string LastRaisedEventName { get; private set; }

        public object[] LastRaisedEventParameters { get; private set; }
    }
}
