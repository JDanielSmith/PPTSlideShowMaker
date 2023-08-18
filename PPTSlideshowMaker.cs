using System;
using System.Collections.Generic;
using System.Diagnostics;
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
	readonly DirectoryInfo directoryInfo;
	readonly IList<string> files = new List<string>();
	string? backgroundMusic;

	const string VideoExtension = @".m4v";
	const string FolderName = @"folder.png"; // for Plex

	static bool IsImage(string path)
	{
		if (!Drawing.IsImage(path))
		{
			return false;
		}

		// Be sure it's not the folder.png file we created in a previous run
		var filename = Path.GetFileName(path);
		return !String.Equals(filename, FolderName, StringComparison.OrdinalIgnoreCase);
	}

	void ProcessFile(FileInfo file)
	{
		var path = Shortcut.Resolve(file.FullName);
		if (IsImage(path))
		{
			files.Add(path);
			return;
		}
		if (file.Extension.Equals(VideoExtension, StringComparison.OrdinalIgnoreCase) ||
			file.Extension.Equals(".json", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}

		// If the file isn't an image and not the video created from a previous run,
		// then it must be background music.
		if (!Drawing.IsImage(path))
		{
			backgroundMusic = path;
		}
	}
	public PPTSlideshowMaker(Settings settings)
	{
		//application.Visible = Office.MsoTriState.msoFalse;
		//application.WindowState = PPT.PpWindowState.ppWindowMinimized;

		presentation = new Presentation(application.Presentations.Add());

		this.settings = settings;
		directoryInfo = new DirectoryInfo(this.settings.Directory!);

		foreach (var file in directoryInfo.EnumerateFiles())
		{
			ProcessFile(file);
		}
	}

	static string Combine(DirectoryInfo rootDirectory, string path)
	{
		return Path.Combine(rootDirectory.FullName, path);
	}

	string? GetBackgroundMusicPath()
	{
		var backgroundMusicPath = settings.BackgroundMusicPath;
		if (backgroundMusicPath is not null)
		{
			if (!File.Exists(backgroundMusicPath))
			{
				backgroundMusicPath = Combine(directoryInfo, backgroundMusicPath);
			}
		}
		else if (backgroundMusic is not null)
		{
			backgroundMusicPath = backgroundMusic;
		}
		return backgroundMusicPath;
	}

	FileInfo? FolderPNG {  get; set; }
	public PPTSlideshowMaker AddTitleSlide()
	{
		var backgroundMusicPath = GetBackgroundMusicPath();
		string title = settings.Title ?? directoryInfo.Name;
		var slide = presentation.AddTitleSlide(title, settings.SubTitle, backgroundMusicPath);

		// "folder.png" for Plex
		FolderPNG = new FileInfo(Combine(directoryInfo, FolderName));
		slide.Export(FolderPNG.FullName, "PNG");

		return this;
	}

	public PPTSlideshowMaker AddPictureSlides()
	{		
		var titleSlideTime = presentation.TitleAdvanceTime + presentation.TransitionDuration;
		var pictureSlidesTime = presentation.MediaLength - titleSlideTime;

		int slides = files.Count();
		var pictureSlideTime = (pictureSlidesTime / slides) - presentation.TransitionDuration;
		presentation.SlideAdvanceTime = (float)pictureSlideTime.TotalSeconds;

		int count = 0; // During development, it can be convenient to stop after just a few pictures
		foreach (var file in files)
		{
			presentation.AddPictureSlide(file);

			count++;
			if (count > slides) break;
		}

		return this;
	}

	public PPTSlideshowMaker AddEndSlide()
	{
		presentation.AddEndSlide(settings.EndTitle, settings.Copyright);
		return this;
	}

	public PPTSlideshowMaker CreateVideo()
	{
		var m4v = directoryInfo.Name + VideoExtension;
		var path = Combine(directoryInfo, m4v);
		presentation.CreateVideo(path);

		if (FolderPNG is not null)
		{
			TagEditor.AddCover(path, FolderPNG!);
		}
		
		return this;
	}

	public void Quit()
	{
		presentation.Close();
		application.Quit();
	}
}
