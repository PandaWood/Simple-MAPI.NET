/******************************************************
Simple MAPI.NET
https://github.com/PandaWood/Simple-MAPI.NET
*******************************************************/

using System;
using Win32Mapi;

namespace SimpleMAPI
{
	class SimpleMapiDemo
	{
		[STAThread]
		static void Main(string[] args)
		{
			if ((args == null) || (args.Length < 3))
			{
				Console.WriteLine("SimpleMAPI Console syntax :\n\tConsole-Demo recip@any.org subject text");
				return;
			}

			SimpleMapi ma = new SimpleMapi();
			//!ma.Logon(IntPtr.Zero)		// this code is strictly correct, but won't work with Outlook 64-bit and is needed in the most common usage

			ma.AddRecipient(args[0], null, false);
			if (!ma.Send(args[1], args[2]))
			{
				Console.WriteLine("MAPI SendMail failed! : " + ma.Error());
				return;
			}

			//ma.Logoff();
			Console.WriteLine("SimpleMAPI Console: email sent successfully.");
		}
	}
}
