static void AddTitleSlide(Presentation presentation, DirectoryInfo directoryInfo, string backgroundMusic,
	string subTitle)
{
	presentation.AddTitleSlide(directoryInfo.Name, subTitle,
		Path.Combine(directoryInfo.FullName, backgroundMusic));
}
static void CreateVideo(Presentation presentation, DirectoryInfo directoryInfo)
{
	var m4v = directoryInfo.Name + ".m4v";
	presentation.CreateVideo(Path.Combine(directoryInfo.FullName, m4v));
}

// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run
var application = new Microsoft.Office.Interop.PowerPoint.Application();
//application.Visible = Office.MsoTriState.msoFalse;
//application.WindowState = PPT.PpWindowState.ppWindowMinimized;
var presentation = new Presentation(application.Presentations.Add());


var rootDirectory = new DirectoryInfo(@"E:\Users\JDani\OneDrive\Archive\Videos\Prague 2016");
AddTitleSlide(presentation, rootDirectory, "10 Good King Wenceslas - Shortcut.lnk",
	"Good King Wenceslas\n(Mannheim Steamroller, with members of the\nCzech Philharmonic Orchestra)");	

presentation.AddPictureSlides(rootDirectory);
presentation.AddEndSlide("Copyright © 2023\nJ. Daniel Smith", DateTime.Now.ToString());

CreateVideo(presentation, rootDirectory);

presentation.Close();
application.Quit();