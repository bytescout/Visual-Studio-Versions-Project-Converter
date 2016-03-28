//*******************************************************************
//       Visual Studio Versions Project Converter		                                     
//                                                                   
//       Copyright © 2016 ByteScout - http://www.bytescout.com       
//       ALL RIGHTS RESERVED                                         
//                                                                   
//*******************************************************************

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VSProjectConverter2
{
	public enum VisualStudioVersion
	{
		VS2005, VS2008, VS2010, VS2012, VS2013, VS2015,
	}

	class Program
	{
		public static string InputFile = null;
		public static List<VisualStudioVersion> TargetVersions = new List<VisualStudioVersion>();
		public static bool RenameCurrent = false;

		static void Main(string[] args)
		{
			if (args.Length == 0)
			{
				PrintUsage();
				
				return;
			}

			foreach (string arg in args)
			{
				if (File.Exists(arg))
				{
					InputFile = Path.GetFullPath(arg);

					string extension = Path.GetExtension(InputFile);
					if (extension != ".csproj" && extension != ".vbproj")
					{
						Console.WriteLine("Unsupported project file type");
						return;
					}
				}
				else if (arg == "/rename")
				{
					RenameCurrent = true;
				}
				else if (arg == "/allversions")
				{
					TargetVersions.AddRange(Enum.GetValues(typeof(VisualStudioVersion)).Cast<VisualStudioVersion>());
				}
				else if (arg == "/vs2005")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2005))
						TargetVersions.Add(VisualStudioVersion.VS2005);
				}
				else if (arg == "/vs2008")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2008))
						TargetVersions.Add(VisualStudioVersion.VS2008);
				}
				else if (arg == "/vs2010")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2010))
						TargetVersions.Add(VisualStudioVersion.VS2010);
				}
				else if (arg == "/vs2012")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2012))
						TargetVersions.Add(VisualStudioVersion.VS2012);
				}
				else if (arg == "/vs2013")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2013))
						TargetVersions.Add(VisualStudioVersion.VS2013);
				}
				else if (arg == "/vs2015")
				{
					if (!TargetVersions.Contains(VisualStudioVersion.VS2015))
						TargetVersions.Add(VisualStudioVersion.VS2015);
				}
			}

			// default is /allversions
			if (TargetVersions.Count == 0)
				TargetVersions.AddRange(Enum.GetValues(typeof(VisualStudioVersion)).Cast<VisualStudioVersion>());

			if (!String.IsNullOrEmpty(InputFile) && TargetVersions.Count > 0)
			{
				try
				{
					using (Converter converter = new Converter(InputFile, TargetVersions, RenameCurrent))
					{
						converter.Go();
					}
				}
				catch (Exception e)
				{
					Console.WriteLine(e);
					Console.WriteLine("Press any key...");
					Console.ReadKey();
				}
			}
		}

		private static void PrintUsage()
		{
			Console.WriteLine("Upgrades or downgrades C# and VB project files to all Visual Studio versions.");
			Console.WriteLine();
			Console.WriteLine("VSProjectConverter2 filename [/addpostfix] [/allversions] [/vs2005] [/vs2008] [/vs2010] [/vs2012] [/vs2013] [/vs2015]");
			Console.WriteLine();
			Console.WriteLine("  /rename       Add \".VS20**\" postfix to input filename, e.g. \"MyProject.VS2010.csproj\".");
			Console.WriteLine("  /allversions  (default) Create projects for all Visual Studio versions.");
			Console.WriteLine("  /vs2005       Create Visual Studio 2005 compatible project.");
			Console.WriteLine("  /vs2008       Create Visual Studio 2008 compatible project.");
			Console.WriteLine("  /vs2010       Create Visual Studio 2010 compatible project.");
			Console.WriteLine("  /vs2012       Create Visual Studio 2012 compatible project.");
			Console.WriteLine("  /vs2013       Create Visual Studio 2013 compatible project.");
			Console.WriteLine("  /vs2015       Create Visual Studio 2015 compatible project.");
			Console.WriteLine();
			Console.WriteLine("Press any key...");
			Console.ReadKey();
		}
	}
}
