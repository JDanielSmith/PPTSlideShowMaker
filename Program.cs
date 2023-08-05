using System.Text.Json;

static string Combine(DirectoryInfo rootDirectory, string path)
{
	return Path.Combine(rootDirectory.FullName, path);
}

static Settings ReadSettings(DirectoryInfo rootDirectory)
{
	// https://learn.microsoft.com/en-us/dotnet/standard/serialization/system-text-json/how-to?pivots=dotnet-8-0
	string fileName = "PPTSlideshowMaker.json";
	string path = Combine(rootDirectory, fileName);
	string jsonString = File.ReadAllText(path);
	return JsonSerializer.Deserialize<Settings>(jsonString)!;

}
static void AddTitleSlide_(Presentation presentation, DirectoryInfo directoryInfo, string backgroundMusic,
	string subTitle)
{
	var backgroundMusicPath = backgroundMusic;
	if (!File.Exists(backgroundMusicPath))
	{
		backgroundMusicPath = Combine(directoryInfo, backgroundMusic);
	}

	presentation.AddTitleSlide(directoryInfo.Name, subTitle, backgroundMusicPath);
}
static void AddTitleSlide(Presentation presentation, DirectoryInfo directoryInfo, Settings settings)
{
	var backgroundMusicPath = settings.BackgroundMusicPath;
	if (backgroundMusicPath is not null)
	{
		if (!File.Exists(backgroundMusicPath))
		{
			backgroundMusicPath = Combine(directoryInfo, backgroundMusicPath);
		}
	}

	string title = settings.Title ?? directoryInfo.Name;
	presentation.AddTitleSlide(title, settings.SubTitle, backgroundMusicPath);
}

static void AddEndSlide(Presentation presentation, Settings settings)
{
	presentation.AddEndSlide(settings.EndTitle, settings.Copyright);
}
static void CreateVideo(Presentation presentation, DirectoryInfo directoryInfo)
{
	var m4v = directoryInfo.Name + ".m4v";
	presentation.CreateVideo(Combine(directoryInfo, m4v));
}

var rootDirectory = new DirectoryInfo(@"C:\Users\JDani\OneDrive\Archive\Videos\Prague 2016");
var settings = ReadSettings(rootDirectory);

// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run
var application = new Microsoft.Office.Interop.PowerPoint.Application();
//application.Visible = Office.MsoTriState.msoFalse;
//application.WindowState = PPT.PpWindowState.ppWindowMinimized;
var presentation = new Presentation(application.Presentations.Add());


//AddTitleSlide(presentation, rootDirectory, "10 Good King Wenceslas - Shortcut.lnk",
//	"Good King Wenceslas\n(Mannheim Steamroller, with members of the\nCzech Philharmonic Orchestra)");	
AddTitleSlide(presentation, rootDirectory, settings);

presentation.AddPictureSlides(rootDirectory);
//presentation.AddEndSlide("Copyright © 2023\nJ. Daniel Smith", DateTime.Now.ToString());
AddEndSlide(presentation, settings);

CreateVideo(presentation, rootDirectory);

presentation.Close();
application.Quit();