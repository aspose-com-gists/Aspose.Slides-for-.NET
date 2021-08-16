# How to Open, Load, or View PowerPoint in C#

To open PowerPoint presentations in C#, you need Aspose.Slides for .NET.

**Aspose.Slides for .NET** is an advanced presentation processing API for PowerPoint that allows applications to generate, read, write, protect, modify, convert, print, and perform other tasks with presentations in .NET C# without using Microsoft PowerPoint or Office. 

**TIP**: You may to check out [Aspose PowerPoint online viewer](https://products.aspose.app/slides/viewer) to see how Aspose used its own APIs to develop an online service that allows people to view presentations for free. 

## Opening PowerPoint in C#

Go through these steps:

1. Create a new C# console application in Visual Studio—or you can load your preferred project. 

2. Install Aspose.Slides through any of these methods:
   * Open NuGet Package Manager, search for *Aspose.Slides*, and then install it. 
   * Go through **Tools** > **Library Package Manager** > **Package Manager Console** and then run this command: `Install-Package Aspose.Slides.NET`

3. Add this namespace to your program.cs file (if it's not there already):

```c#
using Aspose.Slides;
```

4. Use this code snippet to open or load up the PowerPoint:

```c#
// Opens the presentation file by passing the file path to the constructor of Presentation class
Presentation pres = new Presentation("OpenPresentation.pptx");

// Prints the total number of slides present in the presentation
System.Console.WriteLine(pres.Slides.Count.ToString());
```



## TIPS

* [Get Aspose.Slides for .NET.](https://products.aspose.com/slides/net/)
* [Learn more about Aspose.Slides features.](https://docs.aspose.com/slides/net/features-overview/)
* [Read Aspose.Slides documentation.](https://docs.aspose.com/slides/net/) 