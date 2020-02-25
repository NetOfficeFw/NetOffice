# NetOffice Addin in .NET Core 3.1

Sample Word addin running in .NET Core 3.1 using COM hosting model from .NET Core.
This sample supports Microsoft Word 32-bit only.

## Getting Started

Running this sample requires a lot of manual steps so far.
You must run the registration script with administrative privileges.

Registration script will use **regsrv32.exe** to register comhost.dll and
it will create Word Addin registry keys so MS Word can load the add-in.

1. Compile the **SimpleNetCoreAddin** project
2. Register the addin using script **register-addin.cmd** in **01 Simple** folder
