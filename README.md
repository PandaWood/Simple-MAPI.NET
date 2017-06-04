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

Set a subject, body and attachments via a simple (M)API, which then invokes the default email client (eg Outlook/Windows Mail) with the message, ready to send.

## History

I was using this source code being relied on within the [ExceptionReporter.NET](https://github.com/PandaWood/ExceptionReporter.NET) project.

It had been copied, with internal attribution, from the article on [Simple-MAPI-NET](https://www.codeproject.com/Articles/2048/Simple-MAPI-NET) from the Code Project site.

I have made some changes while the code was in there - some semantic, but at least one was a fix to a problem. I'll try and document this shortly.

## How to use

The console demo shows something slightly different, but I have used it, with success, like this (no logon/logoff)
```
using Win32Mapi;

...
var mapi = new Mapi();
mapi.AddRecipient(name: "bob@gmail.com", addr: null, cc: false);
mapi.Attach(filepath: "c:\bob.txt")
mapi.Send(subject: "a subject", noteText: "a body text")
```