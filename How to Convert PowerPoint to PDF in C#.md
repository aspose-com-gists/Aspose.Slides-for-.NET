# How to Convert PowerPoint to PDF in C#

To convert PowerPoint presentations to PDF files in C#, you need to install Aspose.Slides for .NET.

**Aspose.Slides for .NET** is an advanced presentation processing API for PowerPoint that allows applications to generate, read, write, protect, modify, convert, print, and perform other tasks with presentations in .NET C# without using Microsoft PowerPoint or Office. 

**Tip**: You may want to check out the [free PowerPoint to PDF conversion](https://products.aspose.app/slides/conversion/powerpoint-to-pdf) service Aspose developed using its own API.

## Converting PowerPoint to PDF

Go through these steps:

1. Create a new C# console application in Visual Studio. Alternatively, you can load your preferred project. 

2. Install **Aspose.Slides** through any of these methods:
   * Open NuGet Package Manager, search for *Aspose.Slides*, and then install it. 
   * Go through **Tools** > **Library Package Manager** > **Package Manager Console** and then run this command: `Install-Package Aspose.Slides.NET`

3. Add these namespaces to your program.cs file:

```c#
using Aspose.Slides;
using Aspose.Slides.Export;
```

4. Use this code snippet to convert the PowerPoint presentation (PPTX or PPT) to PDF:

```c#
// Instantiates a Presentation object that represents a PPTX file
Presentation presentation = new Presentation("Input-PowerPoint.pptx");

// Saves the presentation as PDF
presentation.Save("Output-PDF.pdf", SaveFormat.Pdf);
```

## Converting PowerPoint to Accessible PDF

If you want to use a PowerPoint to PDF conversion procedure in C# that complies with with [Web Content Accessibility Guidelines (**WCAG**)](https://www.w3.org/TR/WCAG-TECHS/pdf.html), you have to specify the compliance standards or tags: **PDF/A1a**, **PDF/A1b**, and **PDF/UA**.

This sample code shows you how to specify your preferred PDF compliance standard when converting **PPTX to PDF** or **PPT to PDF**:

```c#
using (Presentation pres = new Presentation("Input-Presentation.pptx"))
{
    pres.Save("pres-a1a-compliance.pdf", SaveFormat.Pdf, new PdfOptions()
    {
        Compliance = PdfCompliance.PdfA1a
    });
   
    pres.Save("pres-a1b-compliance.pdf", SaveFormat.Pdf, new PdfOptions()
    {
        Compliance = PdfCompliance.PdfA1b
    });
   
    pres.Save("pres-ua-compliance.pdf", SaveFormat.Pdf, new PdfOptions()
   {
        Compliance = PdfCompliance.PdfUa
    });
}
```



## TIPS

* [Get Aspose.Slides for .NET.](https://products.aspose.com/slides/net/)
* [Learn more about Aspose.Slides features.](https://docs.aspose.com/slides/net/features-overview/)
* [Read Aspose.Slides documentation.](https://docs.aspose.com/slides/net/) 