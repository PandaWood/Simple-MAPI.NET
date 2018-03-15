# Simple-MAPI.NET
A .NET wrapper around Simple MAPI - a set of functions and related data structures to add messaging functionality to Windows-based apps

This project originated from the article [Simple-MAPI-NET](https://www.codeproject.com/Articles/2048/Simple-MAPI-NET) on CodeProject site

### The nuget package
 [![NuGet](https://img.shields.io/nuget/v/Simple-MAPI.NET.svg)](https://www.nuget.org/packages/Simple-MAPI.NET/)
```
PM> Install-Package Simple-MAPI.NET
```

## What does it do?

Basically allows you to setup email messages using the default email client on the user's (Windows-based) machine. 

Set a subject, body and attachments and simple (M)API invokes the default email client (eg Outlook/Windows Mail) with the message, ready to send.

## History

I was using this source code within the [ExceptionReporter.NET](https://github.com/PandaWood/ExceptionReporter.NET) project.

It had been copied, with internal attribution, from the article on [Simple-MAPI-NET](https://www.codeproject.com/Articles/2048/Simple-MAPI-NET) from the Code Project site. But I noticed that no one else had committed this code a to a repository, so me to the rescue. And I wanted my ExceptionReport.NET library to use a dependency rather than horde all this unrelated code.

### Should I use
Simple MAPI is an old technology and it's basically deprecated - Microsoft warns that the [use of Simple MAPI is discouraged and that it may be unavailable in later versions of Windows](https://msdn.microsoft.com/en-us/library/windows/desktop/dd296734(v=vs.85).aspx)

But it's the only way, I know, to invoke the default Windows Email client and set subject/body and attachments.

## How to use

The console demo shows something slightly different, but I have used it, with success, like this (no logon/logoff)
```
using Win32Mapi;

.
var mapi = new SimpleMapi();
mapi.AddRecipient(name: "bob@gmail.com", addr: null, cc: false);
mapi.Attach(filepath: "c:\\bob.txt");
mapi.Send(subject: "a subject", noteText: "a body text");
```

I don't use Logon/off() because at least in one case (64-bit Office/Outlook) it just creates a superflous error but sends anyway.

ie removing the call to Logon() still works.
