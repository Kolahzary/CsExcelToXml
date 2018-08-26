# CsExcelToXml [![Build status](https://ci.appveyor.com/api/projects/status/14g7g3xkp4kw9w1s?svg=true)](https://ci.appveyor.com/project/Kolahzary/csexceltoxml)

Sample project of converting Microsoft Excel files to XML using C#

## How it works?
Simply reads Excel data into memory using [NPOI](https://github.com/tonyqus/npoi) library and then writes it back as XML file using [System.Xml.Serialization.XmlSerializer](https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlserializer?view=netframework-4.7.2).

## Dependancies
- Microsoft Visual Studio (also compilable using cli)
- Microsoft .NET Framework 4.5 (may work on other versions)
- [NPOI Nuget Package](https://www.nuget.org/packages/NPOI/)
