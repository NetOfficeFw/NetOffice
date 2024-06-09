using System;
using System.Collections;
using System.Collections.Generic;
using NetOffice;
using NetOffice.Availability;
using NetOffice.Tools;

/// <summary>
/// This addin using the unit tests is in assembly
/// without the AssemblyTitleAttribute set.
/// </summary>
public class NoAssemblyTitleAddin : COMAddinBase

{
    public override ICOMObject AppInstance => new NoAssemblyTitleAppInstance();

    public override Core Factory => throw new NotImplementedException();

    public override IEnumerable Roots { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
}

public class NoAssemblyTitleAppInstance : ICOMObject
{
    public object SyncRoot => throw new NotImplementedException();

    public Core Factory => throw new NotImplementedException();

    public Invoker Invoker => throw new NotImplementedException();

    public Settings Settings => throw new NotImplementedException();

    public DebugConsole Console => throw new NotImplementedException();

    public object UnderlyingObject => throw new NotImplementedException();

    public Type UnderlyingType => throw new NotImplementedException();

    public string UnderlyingTypeName => throw new NotImplementedException();

    public string UnderlyingFriendlyTypeName => throw new NotImplementedException();

    public string UnderlyingComponentName => throw new NotImplementedException();

    public string InstanceName => "NoAssemblyTitleAppInstance";

    public string InstanceFriendlyName => throw new NotImplementedException();

    public string InstanceComponentName => throw new NotImplementedException();

    public Type InstanceType => throw new NotImplementedException();

    public bool IsDisposed => throw new NotImplementedException();

    public bool IsCurrentlyDisposing => throw new NotImplementedException();

    public ICOMObject ParentObject => throw new NotImplementedException();

    public IEnumerable<ICOMObject> ChildObjects => throw new NotImplementedException();

    public bool IsEventBinding => throw new NotImplementedException();

    public bool IsEventBridgeInitialized => throw new NotImplementedException();

    public bool IsWithEventRecipients => throw new NotImplementedException();

    public event OnDisposeEventHandler OnDispose;

    public void AddChildObject(ICOMObject childObject)
    {
        throw new NotImplementedException();
    }

    public object Clone()
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

    public bool EntityIsAvailable(string name)
    {
        return false;
    }

    public bool EntityIsAvailable(string name, SupportedEntityType searchType)
    {
        return false;
    }

    public bool RemoveChildObject(ICOMObject childObject)
    {
        throw new NotImplementedException();
    }

    T ICOMObject.To<T>()
    {
        throw new NotImplementedException();
    }
}