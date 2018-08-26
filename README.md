# CsExcelToXml [![Build status](https://ci.appveyor.com/api/projects/status/14g7g3xkp4kw9w1s?svg=true)](https://ci.appveyor.com/project/Kolahzary/csexceltoxml)

Sample project of converting Microsoft Excel files to XML using C#

## How it works?
Simply reads Excel data into memory using [NPOI](https://www.nuget.org/packages/NPOI/) library and then writes it back as XML file using [System.Xml.Serialization.XmlSerializer](https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlserializer?view=netframework-4.7.2).
