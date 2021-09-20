using Mono.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChapterResequence
{
	public class Program
	{
		static int verbosity;

		public static void Main(string[] args)
		{
			bool show_help = false;
			bool resequenceListings = false;
			List<string> chapterNumbers = new List<string>();
			List<string> chapterPaths = new List<string>();
			


			var p = new OptionSet() {
			"Usage: [OPTIONS] ",
			"Greet a list of individuals with an optional message.",
			"If no message is specified, a generic greeting is used.",
			"",
			"Inputs:",
			{ "cn|chapter=", "the {Chapter Number} to use\n",
			  chapter => chapterNumbers.Add(chapter) },
			{ "doc|chapterFile=",
				"the {path} to a chapter.docx.\n",
			  chapterPath => chapterPaths.Add(chapterPath)},
			{ "v", "increase debug message verbosity\n",
			  v => { if (v != null) ++verbosity; } },
			"Options:",
			{ "rl|ResequenceListings",  "resequence listings in chapter\n",
			v=>resequenceListings= (v != null) },

			{ "h|help",  "show this message and exit",
			  v => show_help = (v != null) },
		};

			List<string> extra;
			try
			{
				extra = p.Parse(args);
			}
			catch (OptionException e)
			{
				Console.Write("greet: ");
				Console.WriteLine(e.Message);
				Console.WriteLine("Try `-help' for more information.");
				return;
			}

			if (resequenceListings)
			{
				//choose whether to use files or chapter number
				foreach(string chapter in chapterPaths)
				if (File.Exists(chapter))
				{

						Console.WriteLine("Resequencing Listings...");
					//ResequenceTools.ResequenceListings(Path.GetFullPath(chapter), ParseChapterNumber(chapter));
				}

				
			}

			if (show_help || extra.Count>0)
			{
			
				p.WriteOptionDescriptions(Console.Out);
				return;
			}

			

			/*string message;
			if (extra.Count > 0)
			{
				message = string.Join(" ", extra.ToArray());
				Debug("Using new message: {0}", message);
			}
			else
			{
				message = "Hello {0}!";
				Debug("Using default message: {0}", message);
			}*/

			/*foreach (string name in names)
			{
				for (int i = 0; i < repeat; ++i)
					Console.WriteLine(message, name);
			}*/
		}

		static void Debug(string format, params object[] args)
		{
			if (verbosity > 0)
			{
				Console.Write("# ");
				Console.WriteLine(format, args);
			}
		}

		public static int ParseChapterNumber(string chapterFilePath)
		{
			//chapter filename format {Michaelis_Ch09}
			int startofChapterNumber = chapterFilePath.IndexOf("_Ch");

			return int.Parse(chapterFilePath.Substring(startofChapterNumber+3, 2));
		}
	}
}
