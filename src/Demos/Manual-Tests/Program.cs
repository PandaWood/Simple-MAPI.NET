/*
Manual smoke tests for Simple-MAPI.NET.

Windows only, and needs a default mail client configured (Outlook etc).
Nothing is ever actually sent - each test opens a compose window that you
close afterwards.

	dotnet run --project src/Demos/Manual-Tests -- you@example.com

Covers the paths changed in 1.3.0 that can't be tested off Windows. Checks
the harness can make itself are asserted automatically; the rest ask you
what appeared on screen.
*/

using System;
using System.IO;
using Win32Mapi;

namespace ManualTests
{
	static class Program
	{
		static int passed;
		static int failed;

		static int Main(string[] args)
		{
			string to = args.Length > 0 ? args[0] : Ask("Destination email address");
			if (string.IsNullOrEmpty(to))
			{
				Console.WriteLine("An email address is needed - pass it as an argument.");
				return 2;
			}

			Console.WriteLine();
			Console.WriteLine("Simple-MAPI.NET manual tests");
			Console.WriteLine("A compose window opens for most tests. Close it when asked -");
			Console.WriteLine("nothing needs to be sent for the test to be valid.");

			TooManyAttachments();
			BasicSend(to);
			TwoSendsFromOneInstance(to);
			UnicodeSubjectAndBody(to);
			SendWithAttachment(to);
			UserCancels(to);

			Console.WriteLine();
			Console.WriteLine("----------------------------------------");
			Console.WriteLine("passed {0}, failed {1}", passed, failed);
			return failed == 0 ? 0 : 1;
		}

		// Automatic - Send rejects this before it reaches MAPI, so no window opens.
		static void TooManyAttachments()
		{
			Header("0. more than 100 attachments fails instead of silently sending none");

			var mapi = new SimpleMapi();
			for (int i = 0; i < 101; i++)
			{
				mapi.Attach(@"C:\not-a-real-file-" + i + ".txt");
			}

			bool sent = mapi.Send("too many attachments", "body");

			Check("Send returned false", !sent, "returned " + sent);
			Check("error is MAPI_E_TOO_MANY_FILES",
				mapi.ErrorValue == SimpleMapi.RETURN_VALUE.MAPI_E_TOO_MANY_FILES,
				mapi.Error());
		}

		static void BasicSend(string to)
		{
			Header("1. basic send - every send goes through the rewritten SendOnce");

			Console.WriteLine("Expect a compose window with:");
			Console.WriteLine("    To      : " + to);
			Console.WriteLine("    Subject : basic send");
			Console.WriteLine("    Body    : plain body text");
			Pause();

			var mapi = new SimpleMapi();
			mapi.AddRecipient(to, null, false);
			bool sent = mapi.Send("basic send", "plain body text");

			Report(mapi, sent);
			Confirm("Did the window show that To, subject and body?");
		}

		// Reset() moved into a finally as part of the Send rewrite. If it stopped
		// running, the second message would still carry the first one's recipient.
		static void TwoSendsFromOneInstance(string to)
		{
			Header("2. two sends from one instance - state must not leak between them");

			var mapi = new SimpleMapi();

			Console.WriteLine("First window: To " + to + ", subject 'first'.");
			Pause();
			mapi.AddRecipient(to, null, false);
			mapi.Send("first", "first body");

			Console.WriteLine();
			Console.WriteLine("Second window: the address must appear EXACTLY ONCE.");
			Console.WriteLine("Twice means Reset() is no longer clearing state.");
			Pause();
			mapi.AddRecipient(to, null, false);
			mapi.Send("second", "second body");

			Confirm("Did the second window list the recipient exactly once?");
		}

		// On Windows 8+ this takes the MAPISendMailW path. If the client rejects
		// Unicode it should fall back to ANSI and still send, with lossy text.
		static void UnicodeSubjectAndBody(string to)
		{
			Header("3. unicode subject and body");

			const string subject = "\u00dcn\u00efc\u00f6d\u00e9 \u2014 \u65e5\u672c\u8a9e";
			const string body = "accented: \u00e9\u00e8\u00ea\u00eb  greek: \u03b1\u03b2\u03b3  cjk: \u4e2d\u6587";

			Console.WriteLine("Expect non-ASCII text intact in both subject and body.");
			Console.WriteLine("(This console may not render it correctly - the mail client is what counts.)");
			Console.WriteLine("Mangled text, or an error below, means the ANSI fallback was used.");
			Pause();

			var mapi = new SimpleMapi();
			mapi.AddRecipient(to, null, false);
			bool sent = mapi.Send(subject, body);

			Report(mapi, sent);
			Confirm("Did the accented, greek and CJK characters all survive?");
		}

		static void SendWithAttachment(string to)
		{
			Header("4. send with an attachment");

			string path = Path.Combine(Path.GetTempPath(), "simple-mapi-test.txt");
			File.WriteAllText(path, "attachment written by the Simple-MAPI.NET manual test harness");

			Console.WriteLine("Expect one attachment: " + Path.GetFileName(path));
			Pause();

			var mapi = new SimpleMapi();
			mapi.AddRecipient(to, null, false);
			mapi.Attach(path);
			bool sent = mapi.Send("with attachment", "should carry one attachment");

			Report(mapi, sent);
			Confirm("Was the attachment there?");
		}

		// Asserted automatically once you close the window.
		static void UserCancels(string to)
		{
			Header("5. cancelling the compose window reports user abort");

			Console.WriteLine("CLOSE the next window WITHOUT sending.");
			Pause();

			var mapi = new SimpleMapi();
			mapi.AddRecipient(to, null, false);
			bool sent = mapi.Send("cancel me", "close this window without sending");

			Check("Send returned false", !sent, "returned " + sent);
			Check("error is MAPI_E_USER_ABORT",
				mapi.ErrorValue == SimpleMapi.RETURN_VALUE.MAPI_E_USER_ABORT,
				mapi.Error());
		}

		// ---- plumbing ----------------------------------------------------------

		static void Header(string title)
		{
			Console.WriteLine();
			Console.WriteLine("----------------------------------------");
			Console.WriteLine(title);
			Console.WriteLine();
		}

		static void Report(SimpleMapi mapi, bool sent)
		{
			Console.WriteLine("    Send returned {0}, error: {1}", sent, mapi.Error());
		}

		static void Check(string what, bool ok, string detail)
		{
			if (ok)
			{
				passed++;
				Console.WriteLine("    PASS  " + what);
			}
			else
			{
				failed++;
				Console.WriteLine("    FAIL  " + what + (string.IsNullOrEmpty(detail) ? "" : "  (" + detail + ")"));
			}
		}

		static void Confirm(string question)
		{
			Console.Write(question + " [y/n] ");
			string answer = (Console.ReadLine() ?? "").Trim().ToLowerInvariant();
			Check(question, answer.StartsWith("y"), "you answered '" + answer + "'");
		}

		static void Pause()
		{
			Console.Write("Press enter to open the window... ");
			Console.ReadLine();
		}

		static string Ask(string prompt)
		{
			Console.Write(prompt + ": ");
			return (Console.ReadLine() ?? "").Trim();
		}
	}
}
