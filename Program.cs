static Settings ReadSettings(DirectoryInfo rootDirectory)
{
	string fileName = "PPTSlideshowMaker.json";
	string path = Path.Combine(rootDirectory.FullName, fileName);

	// https://learn.microsoft.com/en-us/dotnet/standard/serialization/system-text-json/how-to?pivots=dotnet-8-0
	string jsonString = File.ReadAllText(path);
	var retval = System.Text.Json.JsonSerializer.Deserialize<Settings>(jsonString)!;

	retval.Directory = rootDirectory.FullName;

	return retval;
}

var rootDirectory = @"C:\Users\JDani\OneDrive\Archive\Videos\Prague 2016";
var settings = ReadSettings(new DirectoryInfo(rootDirectory));

var slideshowMaker = new PPTSlideshowMaker(settings)
	.AddTitleSlide()
	.AddPictureSlides()
	.AddEndSlide()
	.CreateVideo();
slideshowMaker.Quit();

