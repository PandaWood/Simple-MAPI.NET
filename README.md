# Simple-MAPI.NET
This project originated from the article
[Simple-MAPI-NET](https://www.codeproject.com/Articles/2048/Simple-MAPI-NET)
on CodeProject site. There has been a few important fixes in this project
since the original was imported.

### The nuget package
 [![NuGet](https://img.shields.io/nuget/v/Simple-MAPI.NET.svg)](https://www.nuget.org/packages/Simple-MAPI.NET/)
```
PM> Install-Package Simple-MAPI.NET
```

## What does it do?

Simple MAPI.NET allows you to create an email message using the user's
default email client on Windows.

You can set a subject, body and attachments and invoke
the default email client (eg Outlook/Windows Mail), ready to send.

It's useful if you want to create an email for the user, without sending
it yourself using SMTP etc.

### Should I use
Simple MAPI is an old technology and it's basically deprecated -
Microsoft warns that the [use of Simple MAPI is discouraged and that it may be unavailable in later versions of Windows](https://msdn.microsoft.com/en-us/library/windows/desktop/dd296734(v=vs.85).aspx)

But it's the only way, I know, to invoke the default Email
client and set subject/body and attachments on Windows without using interops or anything Office-specific. Simple MAPI is just a set of functions and related data structures, in Windows, that this library calls.

It still works on Windows 10, with Outlook as the email client.

## How to use

I have used it, with success, like this:
```
using Win32Mapi;

.
var mapi = new SimpleMapi();
mapi.AddRecipient(name: "bob@gmail.com", addr: null, cc: false);
mapi.Attach(filepath: "c:\\bob.txt");
mapi.Send(subject: "a subject", noteText: "a body text");
```

I don't use Logon/off() because it used to cause issues (only in 64-bit Office/Outlook)
and showed a superflous error but sent anyway - but I may need to revisit that choice.

## History

I was using this source code within the [ExceptionReporter.NET](https://github.com/PandaWood/ExceptionReporter.NET) project.

I copied it, with internal attribution, from the article on
[Simple-MAPI-NET](https://www.codeproject.com/Articles/2048/Simple-MAPI-NET)
on the Code Project site.
But I noticed that no one else had committed this code a to a repository...
so, *me* to the rescue.

This way, I could use the library as a dependency rather than
horde a load of unrelated code in the project - and get fixes and suggestions
from the community!
