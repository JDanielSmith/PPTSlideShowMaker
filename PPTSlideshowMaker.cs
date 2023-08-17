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
	readonly DirectoryInfo directoryInfo;
	readonly IList<string> files = new List<string>();
	readonly string? backgroundMusic;

	static readonly string VideoExtension = ".m4v";
	public PPTSlideshowMaker(Settings settings)
	{
		//application.Visible = Office.MsoTriState.msoFalse;
		//application.WindowState = PPT.PpWindowState.ppWindowMinimized;

		presentation = new Presentation(application.Presentations.Add());

		this.settings = settings;
		directoryInfo = new DirectoryInfo(this.settings.Directory!);

		foreach (var file in directoryInfo.EnumerateFiles())
		{
			var path = Shortcut.Resolve(file.FullName);
			if (Drawing.IsImage(path))
			{
				files.Add(path);
			}
			else
			{
				if (file.Extension.Equals(VideoExtension, StringComparison.OrdinalIgnoreCase) ||
					file.Extension.Equals(".json", StringComparison.OrdinalIgnoreCase))
				{
					/* do nothing */
				}
				else
				{
					// If the file isn't an image and not the video created from a previous run,
					// then it must be background music.
					backgroundMusic = path;
				}
			}	
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

	public PPTSlideshowMaker AddTitleSlide()
	{
		var backgroundMusicPath = GetBackgroundMusicPath();
		string title = settings.Title ?? directoryInfo.Name;
		presentation.AddTitleSlide(title, settings.SubTitle, backgroundMusicPath);

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
		presentation.CreateVideo(Combine(directoryInfo, m4v));
		return this;
	}

	public void Quit()
	{
		presentation.Close();
		application.Quit();
	}
}
