using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

internal class PPTSlideshowMaker
{
	// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run
	readonly PPT.Application application = new PPT.Application();
	readonly Presentation presentation;
	readonly Settings settings;
	public PPTSlideshowMaker(Settings settings)
	{
		//application.Visible = Office.MsoTriState.msoFalse;
		//application.WindowState = PPT.PpWindowState.ppWindowMinimized;

		presentation = new Presentation(application.Presentations.Add());

		this.settings = settings;
	}
}
