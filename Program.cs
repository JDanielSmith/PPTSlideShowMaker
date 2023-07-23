using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PPTSlideshowMaker
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run

			var application = new PPT.Application();
			//application.Visible = Office.MsoTriState.msoFalse;
			application.WindowState = PPT.PpWindowState.ppWindowMinimized;
            var presentations = application.Presentations;
			var presentation = new Presentation(presentations.Add());

			presentation.AddTitleSlide("Prague 2016",
				"Good King Wenceslas\n(Mannheim Steamroller, with members of the Czech Philharmonic Orchestra)",
                @"C:\Users\JDani\Music\iTunes\iTunes Media\Music\Mannheim Steamroller\Christmas Symphony II\10 Good King Wenceslas.m4a");
			presentation.TitleAdvanceTime = TimeSpan.FromSeconds(5.0);

			string root = @"C:\Users\JDani\OneDrive\Archive\Pictures\Honeymoon 😘😘😘";
            presentation.AddPictureSlides(Path.Combine(root, "Prague"));

			presentation.CreateVideo(Path.Combine(root, "Prague 2016", "Prague 2016.m4v"));

            presentation.Close();
			application.Quit();
        }
	}
}