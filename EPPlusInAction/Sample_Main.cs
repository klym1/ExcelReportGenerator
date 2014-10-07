//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.IO;
//
//namespace EPPlusInAction
//{
//	class Sample_Main
//	{
//		static void Main(string[] args)
//		{
//			try
//			{
//				DirectoryInfo outputDir = new DirectoryInfo(@"C:\Users\Debopam\Desktop");
//				if (!outputDir.Exists) throw new Exception("Output Directory doesn't exists");
//				string output = "";
//
//				// Executing Sample1
//				Console.WriteLine("Running Sample 1");
//				output = Sample1.RunSample1(outputDir);
//				Console.WriteLine("Sample1 created: {0}", output);
//				Console.WriteLine();
//
//				// Executing Sample2
//				Console.WriteLine("Running Sample 2");
//				output = Sample2.RunSample2(outputDir);
//				Console.WriteLine("Sample2 created: {0}", output);
//				Console.WriteLine();
//			}
//			catch (Exception ex)
//			{
//				Console.WriteLine("Error: {0}", ex.Message);
//			}
//			Console.WriteLine();
//			Console.WriteLine("Press any key to exit...");
//			Console.Read();
//		}
//	}
//}
