using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace SlideshowMaker
{
	internal class Program
	{

		static Shape addPictureSlide(PPT.Presentation presentation, string filename)
		{
			var master = presentation.SlideMaster;
			var pptLayout = master.CustomLayouts[1];

			var index = presentation.Slides.Count + 1;
			var pptSlide = presentation.Slides.AddSlide(index, pptLayout);
			while (pptSlide.Shapes.Count > 0)
			{
				pptSlide.Shapes[1].Delete();
			}
			pptSlide.ColorScheme[PPT.PpColorSchemeIndex.ppBackground].RGB = Int32.MinValue; // Black

			var shape = pptSlide.Shapes.AddPicture(filename, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 0, 0);
			shape.Left = (master.Width - shape.Width) / 2;
			return shape;
		}

		static void Main(string[] args)
		{
			// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run

			var application = new PPT.Application();
			var presentations = application.Presentations;
			var presentation = presentations.Add();

			var fn = @"E:\Users\JDani\OneDrive\Archive\Pictures\Honeymoon 😘😘😘\Prague\20161229_083820886_iOS.jpg";
			addPictureSlide(presentation, fn);

			//presentation.Close();
			//application.Quit();
		}
	}
}