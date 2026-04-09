using System;
using System.Collections.Generic;
using NetOffice;
using NetOffice.Exceptions;

namespace NetOffice.Benchmarks
{
    /// <summary>
    /// Mock implementation of ICOMObject for benchmarking purposes.
    /// This is a minimal implementation focusing on the properties needed for collection operations.
    /// </summary>
    internal class MockCOMObject : ICOMObject
    {
        private static int _instanceCounter = 0;
        private readonly int _id;
        private bool _isDisposed;

        public MockCOMObject()
        {
            _id = System.Threading.Interlocked.Increment(ref _instanceCounter);
        }

        // ICOMObjectProxy implementation
        public object UnderlyingObject => new object();
        public Type UnderlyingType => typeof(object);
        public string UnderlyingTypeName => "MockCOMObject";
        public string UnderlyingFriendlyTypeName => "MockCOMObject";
        public string UnderlyingComponentName => "NetOffice.Benchmarks";
        public string InstanceName => $"MockCOMObject_{_id}";
        public string InstanceFriendlyName => $"MockCOMObject_{_id}";
        public string InstanceComponentName => "NetOffice.Benchmarks";
        public Type InstanceType => typeof(MockCOMObject);

        // ICOMObjectDisposable implementation
        public event OnDisposeEventHandler? OnDispose;
        public bool IsDisposed => _isDisposed;
        public bool IsCurrentlyDisposing { get; private set; }

        public void Dispose()
        {
            Dispose(true);
        }

        public void Dispose(bool disposeEventBinding)
        {
            if (_isDisposed || IsCurrentlyDisposing)
                return;

            IsCurrentlyDisposing = true;
            try
            {
                // OnDisposeEventArgs constructor is internal, so we can't invoke the event properly
                // For benchmarking purposes, we'll skip event invocation
                _isDisposed = true;
            }
            finally
            {
                IsCurrentlyDisposing = false;
            }
        }

        // ICOMObject implementation
        public object SyncRoot => this;
        public Core? Factory { get; set; }
        public Invoker? Invoker { get; set; }
        public Settings? Settings { get; set; }
        public DebugConsole? Console { get; set; }

        public T To<T>() where T : class, ICOMObject
        {
            throw new NotImplementedException("MockCOMObject does not support conversion");
        }

        public object Clone()
        {
            return new MockCOMObject();
        }

        // ICOMObjectTable implementation
        public ICOMObject? ParentObject { get; set; }
        public IEnumerable<ICOMObject> ChildObjects => Array.Empty<ICOMObject>();

        public void AddChildObject(ICOMObject childObject)
        {
            // No-op for mock
        }

        public bool RemoveChildObject(ICOMObject childObject)
        {
            return false; // No-op for mock
        }

        // ICOMObjectTableDisposable implementation
        public void DisposeChildInstances()
        {
            // No-op for mock
        }

        public void DisposeChildInstances(bool disposeEventBinding)
        {
            // No-op for mock
        }

        // ICOMObjectEvents implementation
        public bool IsEventBinding => false;
        public bool IsEventBridgeInitialized => false;
        public bool IsWithEventRecipients => false;

        // ICOMObjectAvailability implementation
        public bool EntityIsAvailable(string name)
        {
            return true;
        }

        public bool EntityIsAvailable(string name, Availability.SupportedEntityType searchType)
        {
            return true;
        }

        // Override GetHashCode and Equals for use in HashSet
        public override int GetHashCode()
        {
            return _id.GetHashCode();
        }

        public override bool Equals(object? obj)
        {
            if (obj is MockCOMObject other)
            {
                return _id == other._id;
            }
            return false;
        }
    }
}
