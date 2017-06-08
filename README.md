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

.
var mapi = new Mapi();
mapi.AddRecipient(name: "bob@gmail.com", addr: null, cc: false);
mapi.Attach(filepath: "c:\bob.txt")
mapi.Send(subject: "a subject", noteText: "a body text")
```

I don't use Logon/off() because at least in one case (64-bit Office/Outlook) it just creates a superflous error but sends anyway.

ie removing the call to Logon() still works.

Here's the initial issue writeup:

> I'm trying to use ExceptionReporter in a managed Win32 Application.  
> I'm running on Win7/64, with 64-bit Office (including Outlook) installed and no 32-bit mail client (Win7 does not come with any mail client by default).
> When I try to get ER to send an Email, I see an error "Microsoft Office Outlook / Either there is no default mail client or the current mail client cannot fulfill the messaging request. Please run Microsoft Outlook and set it as the default mail client". 
> It turns out that Outlook is the default mail client.  
> The message box is shown during the (first) call to MAPILogon in Mapi.Logon(), and the error value returned is 0x80004005; session is left at null.  However, the code continues and sends the email successfully.  Looking at the documentation for MAPILogon here: [MAPILogon function](http://msdn.microsoft.com/en-us/library/windows/desktop/dd296726(v=vs.85).aspx)
> it seems that the function is deprecated.
> I've disabled the call to mapi.Logon in MailSender.SendMapi and everything still seems to work fine.  Since the call seems to be superfluous and troublesome, perhaps it would be better removed from the code?
