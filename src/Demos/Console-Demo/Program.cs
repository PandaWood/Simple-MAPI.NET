/******************************************************
Simple MAPI.NET
https://github.com/PandaWood/Simple-MAPI.NET
*******************************************************/

using System;

namespace SimpleMapi.Demo
{
	class SimpleMapiDemo
	{
		[STAThread]
		static void Main(string[] args)
		{
			if (args == null || args.Length < 3)
			{
				Console.WriteLine("SimpleMAPI Console syntax: ");
				Console.WriteLine("\tSimpleMapi-Demo [email] [subject] [body] [file to attach]");
				Console.WriteLine("\teg SimpleMapi-Demo test@gmail.com 'the subject' 'body' 'c:/test.log'");
				return;
			}

			var simpleMapi = new Win32Mapi.SimpleMapi();
			simpleMapi.AddRecipient(args[2], null, false);

			if (args.Length > 3)
			{
				simpleMapi.Attach(args[3]);
			}

			if (!simpleMapi.Send(args[1], args[2]))
			{
				Console.WriteLine("MAPI SendMail failed: " + simpleMapi.Error());
				return;
			}
		
			Console.WriteLine("SimpleMAPI Console: email sent.");
		}
	}
}
