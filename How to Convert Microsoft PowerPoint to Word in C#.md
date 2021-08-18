# How to Convert PowerPoint to Word in C#

To convert a PowerPoint file (PPTX or PPT) to Word (DOCX or DOCX), you need Aspose.Slides and Aspose.Words for .NET.

* **Aspose.Slides for .NET** is an advanced presentation processing API for PowerPoint that allows applications to generate, read, write, protect, modify, convert, print, and perform other tasks with presentations in .NET C# without using Microsoft PowerPoint. 
* **Aspose.Words** is an advanced document processing API that allows applications to generate, modify, convert, render, print files, and perform other tasks with document without utilizing Microsoft Word. 

**INFO**: The PowerPoint to Word conversion process has been implemented in Aspose [FREE online PowerPoint to Word converter](https://products.aspose.app/slides/conversion/pptx-to-docx). 

## Converting PowerPoint to Word 

Go through these steps:

1. Create a new C# console application in Visual Studio—or you can load your preferred project. 

2. Install Aspose.Slides for .NET and Aspose.Words through any of these methods:
   * Open NuGet Package Manager, search for *Aspose.Slides.NET*, and then install it. Search for *Aspose.Words* and install it too. 
   * Go through **Tools** > **Library Package Manager** > **Package Manager Console** and then run these commands:
     *  `Install-Package Aspose.Slides.NET` 
     * `Install-Package Aspose.Words`
   
3. Add these namespaces to your program.cs file:

```c#
using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Slides;
using Aspose.Words;
using SkiaSharp;
```

4. Use this code snippet to convert the PowerPoint to Word:

```c#
using var presentation = new Presentation();
var doc = new Document();
var builder = new DocumentBuilder(doc);
foreach (var slide in presentation.Slides)
{
   // generates and inserts slide image
   using var bitmap = slide.GetThumbnail(1, 1);
   using var stream = new MemoryStream();
   bitmap.Save(stream, ImageFormat.Png);
   stream.Seek(0, SeekOrigin.Begin);
   using var skBitmap = SKBitmap.Decode(stream);
   builder.InsertImage(skBitmap);

   // inserts slide's texts
   foreach (var shape in slide.Shapes)
   {
      if (shape is AutoShape autoShape)
      {
         builder.Writeln(autoShape.TextFrame.Text);
      }
   }

   builder.InsertBreak(BreakType.PageBreak);
}
```



## TIPS

* [Get Aspose.Slides for .NET.](https://products.aspose.com/slides/net/)
* [Learn more about Aspose.Slides features.](https://docs.aspose.com/slides/net/features-overview/)
* [Read Aspose.Slides documentation.](https://docs.aspose.com/slides/net/) 