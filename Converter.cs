using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace VSProjectConverter2
{
	public class Converter : IDisposable
	{
		public string InputFile { get; set; }
		public List<VisualStudioVersion> TargetVersions { get; set; }
		public bool RenameCurrent { get; set; }

		public Converter(string inputFile, List<VisualStudioVersion> targetVersions, bool renameCurrent)
		{
			InputFile = inputFile;
			TargetVersions = targetVersions;
			RenameCurrent = renameCurrent;
		}

		public void Dispose()
		{
		}

		public void Go()
		{
			VisualStudioVersion currentVSVersion;

			XmlDocument document = new XmlDocument();
			XmlNamespaceManager namespaceManager = new XmlNamespaceManager(new NameTable());

			namespaceManager.AddNamespace("msbuild", "http://schemas.microsoft.com/developer/msbuild/2003");

			string strProjectXML = File.ReadAllText(InputFile);
			
			document.LoadXml(strProjectXML);

			XmlNode projectNode = document.SelectSingleNode("/msbuild:Project", namespaceManager);
			XmlNode toolsVersionNode = document.SelectSingleNode("/msbuild:Project/@ToolsVersion", namespaceManager);

			if (projectNode == null)
				throw new ApplicationException("Invalid project file.");

			if (toolsVersionNode != null)
			{
				switch (toolsVersionNode.InnerText)
				{
					case "3.5":
						currentVSVersion = VisualStudioVersion.VS2008;
						break;
					case "4.0":
						currentVSVersion = VisualStudioVersion.VS2010; // or 2012
						if (strProjectXML.Contains("<Import Project=\"$(MSBuildExtensionsPath)"))
							currentVSVersion = VisualStudioVersion.VS2012;
						break;
					case "12.0":
						currentVSVersion = VisualStudioVersion.VS2013;
						break;
					case "14.0":
						currentVSVersion = VisualStudioVersion.VS2015;
						break;
					default:
						throw new ApplicationException("Unknow Visual Studio version.");
				}
			}
			else
				currentVSVersion = VisualStudioVersion.VS2005;

			foreach (VisualStudioVersion version in TargetVersions)
			{
				if (version == currentVSVersion)
				{
					if (RenameCurrent)
					{
						string newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + 
							"." + currentVSVersion.ToString() + Path.GetExtension(InputFile));
						File.Move(InputFile, newFileName);
					}
				}
				else
				{
					Convert(strProjectXML, currentVSVersion, version);
				}
			}

		}

		void Convert(string strProjectXML, VisualStudioVersion fromVersion, VisualStudioVersion toVersion)
		{
			XmlDocument xmlDocument = new XmlDocument();
			string newFileName = null;

			xmlDocument.LoadXml(strProjectXML);
			
			XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
			namespaceManager.AddNamespace("msbuild", "http://schemas.microsoft.com/developer/msbuild/2003");

			// <Project>
			XmlNode projectNode = xmlDocument.SelectSingleNode("/msbuild:Project", namespaceManager);
			if (projectNode == null)
				throw new ApplicationException("Invalid project file.");

			// <Project ToolsVersion> 
			XmlAttribute toolsVersionAttribute = projectNode.Attributes["ToolsVersion"];

			// <Project><PropertyGroup>
			XmlNode projectPropertyGroupNode = projectNode.SelectSingleNode("msbuild:PropertyGroup", namespaceManager);

			// <Project><PropertyGroup><ProductVersion>
			XmlNode productVersionNode = projectPropertyGroupNode.SelectSingleNode("msbuild:ProductVersion", namespaceManager);
			
			// <Project><PropertyGroup><TargetFrameworkVersion>
			XmlNode targetFrameworkNode = projectPropertyGroupNode.SelectSingleNode("msbuild:TargetFrameworkVersion", namespaceManager);

			// <Project><ItemGroup><Reference>
			XmlNodeList referenceNodeList = projectNode.SelectNodes("msbuild:ItemGroup/msbuild:Reference", namespaceManager);
			
			// <Project><ItemGroup><Import>
			XmlNodeList vbImports = projectNode.SelectNodes("msbuild:ItemGroup/msbuild:Import", namespaceManager);

			// <Project><Import Project> 
			XmlNodeList importProjectAttributes = projectNode.SelectNodes("msbuild:Import/@Project", namespaceManager);


			if (toVersion == VisualStudioVersion.VS2005)
			{
				// Remove <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					projectNode.Attributes.Remove(toolsVersionAttribute);

				// Set <ProductVersion>
				if (productVersionNode != null)
					productVersionNode.InnerText = "8.0.50727";
				else
				{
					productVersionNode = xmlDocument.CreateElement("ProductVersion", namespaceManager.LookupNamespace("msbuild"));
					productVersionNode.InnerText = "8.0.50727";
					projectPropertyGroupNode.AppendChild(productVersionNode);
				}

				// Remove <TargetFrameworkVersion>
				if (targetFrameworkNode != null) 
					projectPropertyGroupNode.RemoveChild(targetFrameworkNode);

				// Remove invalid references
				RemoveReference(referenceNodeList, "Microsoft.CSharp");
				RemoveReference(referenceNodeList, "System.Net.Http");
				RemoveReference(vbImports, "System.Threading.Tasks");

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildToolsPath"))
						attribute.Value = attribute.Value.Replace("MSBuildToolsPath", "MSBuildBinPath");	
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2005" + Path.GetExtension(InputFile));
			}
			else if (toVersion == VisualStudioVersion.VS2008)
			{
				// Set <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					toolsVersionAttribute.Value = "3.5";
				else
				{
					toolsVersionAttribute = xmlDocument.CreateAttribute("ToolsVersion");
					toolsVersionAttribute.Value = "3.5";
					projectNode.Attributes.InsertAfter(toolsVersionAttribute, null);
				}

				// Set <ProductVersion>
				if (productVersionNode != null)
					productVersionNode.InnerText = "9.0.30729";
				else
				{
					productVersionNode = xmlDocument.CreateElement("ProductVersion", xmlDocument.DocumentElement.NamespaceURI);
					productVersionNode.InnerText = "9.0.30729";
					projectPropertyGroupNode.AppendChild(productVersionNode);
				}
				
				// Tune <TargetFrameworkVersion>
				if (targetFrameworkNode != null)
				{
					string versionString = targetFrameworkNode.InnerText;
					int majorVersion = int.Parse(versionString.Substring(1, 1));
					if (majorVersion > 3)
						targetFrameworkNode.InnerText = "v3.5";
				}
				else
				{
					targetFrameworkNode = xmlDocument.CreateElement("TargetFrameworkVersion", xmlDocument.DocumentElement.NamespaceURI);
					targetFrameworkNode.InnerText = "v3.5";
					projectPropertyGroupNode.AppendChild(targetFrameworkNode);
				}

				// Remove invalid references
				RemoveReference(referenceNodeList, "Microsoft.CSharp");
				RemoveReference(referenceNodeList, "System.Net.Http");
				RemoveReference(vbImports, "System.Threading.Tasks");

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildBinPath"))
						attribute.Value = attribute.Value.Replace("MSBuildBinPath", "MSBuildToolsPath");
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2008" + Path.GetExtension(InputFile));
			}
			else if (toVersion == VisualStudioVersion.VS2010)
			{
				// Set <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					toolsVersionAttribute.Value = "4.0";
				else
				{
					toolsVersionAttribute = xmlDocument.CreateAttribute("ToolsVersion");
					toolsVersionAttribute.Value = "4.0";
					projectNode.Attributes.InsertAfter(toolsVersionAttribute, null);
				}

				// Remove <ProductVersion>
				if (productVersionNode != null)
					projectPropertyGroupNode.RemoveChild(productVersionNode);

				// Tune <TargetFrameworkVersion>
				if (targetFrameworkNode != null)
				{
					string versionString = targetFrameworkNode.InnerText;
					int majorVersion = int.Parse(versionString.Substring(1, 1));
					int minorVersion = int.Parse(versionString.Substring(3, 1));
					if (majorVersion >= 4 && minorVersion > 0)
						targetFrameworkNode.InnerText = "v4.0";
				}
				else
				{
					targetFrameworkNode = xmlDocument.CreateElement("TargetFrameworkVersion", namespaceManager.LookupNamespace("msbuild"));
					targetFrameworkNode.InnerText = "v4.0";
					projectPropertyGroupNode.AppendChild(targetFrameworkNode);
				}

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildBinPath"))
						attribute.Value = attribute.Value.Replace("MSBuildBinPath", "MSBuildToolsPath");
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2010" + Path.GetExtension(InputFile));
			}
			else if (toVersion == VisualStudioVersion.VS2012)
			{
				// Set <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					toolsVersionAttribute.Value = "4.0";
				else
				{
					toolsVersionAttribute = xmlDocument.CreateAttribute("ToolsVersion");
					toolsVersionAttribute.Value = "4.0";
					projectNode.Attributes.InsertAfter(toolsVersionAttribute, null);
				}

				// Remove <ProductVersion>
				if (productVersionNode != null)
					projectPropertyGroupNode.RemoveChild(productVersionNode);

				// Tune <TargetFrameworkVersion>
				if (targetFrameworkNode != null)
				{
					string versionString = targetFrameworkNode.InnerText;
					int majorVersion = int.Parse(versionString.Substring(1, 1));
					int minorVersion = int.Parse(versionString.Substring(3, 1));
					if (majorVersion >= 4 && minorVersion > 5)
						targetFrameworkNode.InnerText = "v4.5";
				}
				else
				{
					targetFrameworkNode = xmlDocument.CreateElement("TargetFrameworkVersion", namespaceManager.LookupNamespace("msbuild"));
					targetFrameworkNode.InnerText = "v4.0";
					projectPropertyGroupNode.AppendChild(targetFrameworkNode);
				}

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildBinPath"))
						attribute.Value = attribute.Value.Replace("MSBuildBinPath", "MSBuildToolsPath");
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2012" + Path.GetExtension(InputFile));
			}
			else if (toVersion == VisualStudioVersion.VS2013)
			{
				// Set <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					toolsVersionAttribute.Value = "12.0";
				else
				{
					toolsVersionAttribute = xmlDocument.CreateAttribute("ToolsVersion");
					toolsVersionAttribute.Value = "12.0";
					projectNode.Attributes.InsertAfter(toolsVersionAttribute, null);
				}

				// Remove <ProductVersion>
				if (productVersionNode != null)
					projectPropertyGroupNode.RemoveChild(productVersionNode);

				// Tune <TargetFrameworkVersion>
				if (targetFrameworkNode != null)
				{
					string versionString = targetFrameworkNode.InnerText;
					int majorVersion = int.Parse(versionString.Substring(1, 1));
					int minorVersion = int.Parse(versionString.Substring(3, 1));
					if (majorVersion >= 4 && minorVersion > 5)
						targetFrameworkNode.InnerText = "v4.5";
				}
				else
				{
					targetFrameworkNode = xmlDocument.CreateElement("TargetFrameworkVersion", namespaceManager.LookupNamespace("msbuild"));
					targetFrameworkNode.InnerText = "v4.5";
					projectPropertyGroupNode.AppendChild(targetFrameworkNode);
				}

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildBinPath"))
						attribute.Value = attribute.Value.Replace("MSBuildBinPath", "MSBuildToolsPath");
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2013" + Path.GetExtension(InputFile));
			}
			else if (toVersion == VisualStudioVersion.VS2015)
			{
				// Set <Project ToolsVersion> attribute
				if (toolsVersionAttribute != null)
					toolsVersionAttribute.Value = "14.0";
				else
				{
					toolsVersionAttribute = xmlDocument.CreateAttribute("ToolsVersion");
					toolsVersionAttribute.Value = "14.0";
					projectNode.Attributes.InsertAfter(toolsVersionAttribute, null);
				}

				// Remove <ProductVersion>
				if (productVersionNode != null)
					projectPropertyGroupNode.RemoveChild(productVersionNode);

				// Tune <TargetFrameworkVersion>
				if (targetFrameworkNode != null)
				{
					string versionString = targetFrameworkNode.InnerText;
					int majorVersion = int.Parse(versionString.Substring(1, 1));
					int minorVersion = int.Parse(versionString.Substring(3, 1));
					if (majorVersion >= 4 && minorVersion > 6)
						targetFrameworkNode.InnerText = "v4.6";
				}
				else
				{
					targetFrameworkNode = xmlDocument.CreateElement("TargetFrameworkVersion", namespaceManager.LookupNamespace("msbuild"));
					targetFrameworkNode.InnerText = "v4.5";
					projectPropertyGroupNode.AppendChild(targetFrameworkNode);
				}

				// Fix <Import Project>
				foreach (XmlNode attribute in importProjectAttributes)
				{
					if (attribute.Value.Contains("MSBuildBinPath"))
						attribute.Value = attribute.Value.Replace("MSBuildBinPath", "MSBuildToolsPath");
				}

				// Make file name
				newFileName = Path.Combine(Path.GetDirectoryName(InputFile), Path.GetFileNameWithoutExtension(InputFile) + ".VS2015" + Path.GetExtension(InputFile));
			}

			if (newFileName != null) 
				xmlDocument.Save(newFileName);
		}

		private void RemoveReference(XmlNodeList referenceNodeList, string referenceName)
		{
			foreach (XmlNode referenceNode in referenceNodeList)
			{
				if (referenceNode.Attributes["Include"].Value == referenceName)
				{
					referenceNode.ParentNode.RemoveChild(referenceNode);
					return;
				}
			}
		}
	}
}
