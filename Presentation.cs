using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

internal class Presentation
{
	readonly PPT.Presentation presentation;
	readonly PPT.Master master;
	readonly PPT.CustomLayout layout;
	public Presentation(PPT.Presentation presentation)
	{
		this.presentation = presentation;
		master = this.presentation.SlideMaster;
		layout = master.CustomLayouts[1];
	}

	public TimeSpan TransitionDuration { get; init; } = TimeSpan.FromSeconds(1);

	static void DeleteAllShapes(PPT.Slide slide)
	{
		while (slide.Shapes.Count > 0)
		{
			slide.Shapes[1].Delete();
		}
	}

	public PPT.PpEntryEffect TransitionEntryEffect { get; init; } = PPT.PpEntryEffect.ppEffectRandom;

	PPT.Slide AddSlide(int index = -1)
	{
		if (index < 0)
		{
			index = presentation.Slides.Count + 1;
		}
		var slide = presentation.Slides.AddSlide(index, layout);
		slide.ColorScheme[PPT.PpColorSchemeIndex.ppBackground].RGB = Int32.MinValue; // Black

		var transition = slide.SlideShowTransition;
		transition.EntryEffect = TransitionEntryEffect;
		transition.Duration = (float) TransitionDuration.TotalSeconds;
		transition.AdvanceOnClick = Office.MsoTriState.msoFalse;
		transition.AdvanceOnTime = Office.MsoTriState.msoTrue;

		return slide;
	}

	static void SetText(PPT.Shape shape, string text)
	{
		var textRange = shape.TextFrame.TextRange;
		textRange.Text = text;
		textRange.Font.Color.RGB = Int32.MaxValue; // White
	}

	public TimeSpan TitleAdvanceTime { get; init; } = TimeSpan.FromSeconds(4.0);

	TimeSpan mediaLegnth = TimeSpan.Zero;

	static string GetTargetPath(string lnkPath)
	{
		var shell = new IWshRuntimeLibrary.WshShell();
		dynamic shortcut = shell.CreateShortcut(lnkPath);
		return shortcut.TargetPath;
	}

	static string ResolveShortcut_(string path)
	{
		var extension = System.IO.Path.GetExtension(path);
		if (String.Equals(extension, ".lnk", StringComparison.InvariantCultureIgnoreCase))
		{
			return GetTargetPath(path);
		}
		return path;
	}
	static string ResolveShortcut(string path)
	{
		var retval = ResolveShortcut_(path);
		if (!File.Exists(retval))
		{
			throw new FileNotFoundException(retval);
		}
		return retval;
	}

	public PPT.Slide AddTitleSlide(string title, string subTitle, string backgroundMusicPathname)
	{
		var slide = AddSlide();
		slide.SlideShowTransition.AdvanceTime = (float)TitleAdvanceTime.TotalSeconds;

		SetText(slide.Shapes[1], title);
		SetText(slide.Shapes[2], subTitle);


		var linkToFile = Office.MsoTriState.msoTrue;
		var saveWithDocument = Office.MsoTriState.msoFalse;
		var shape = slide.Shapes.AddMediaObject2(ResolveShortcut(backgroundMusicPathname), linkToFile, saveWithDocument);

		mediaLegnth = TimeSpan.FromMilliseconds(shape.MediaFormat.Length);

		// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.playsettings
		var animationSettings = shape.AnimationSettings;
		animationSettings.AdvanceMode = PPT.PpAdvanceMode.ppAdvanceOnTime;
		var playSettings = animationSettings.PlaySettings;
		playSettings.PlayOnEntry = Office.MsoTriState.msoTrue;
		playSettings.PauseAnimation = Office.MsoTriState.msoFalse;
		playSettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
		playSettings.StopAfterSlides = Int32.MaxValue;

		return slide;
	}


	static PPT.Shape AddPicture(PPT.Shapes shapes, string fileName)
	{
		fileName = ResolveShortcut(fileName);
		try
		{
			using var image = System.Drawing.Image.FromFile(fileName);
		}
		catch (OutOfMemoryException)		
		{
			return null;
		}

		var linkToFile = Office.MsoTriState.msoTrue;
		var saveWithDocument = Office.MsoTriState.msoFalse;
		try
		{
			return shapes.AddPicture(ResolveShortcut(fileName), linkToFile, saveWithDocument, Left: 0, Top: 0);
		}
		catch (System.Runtime.InteropServices.COMException)
		{
			return null;
		}
	}

	public PPT.Slide AddPictureSlide(string filename, int index = -1)
	{
		var slide = AddSlide();
		DeleteAllShapes(slide);
		slide.SlideShowTransition.AdvanceTime = slideAdvanceTime;

		var shape = AddPicture(slide.Shapes, filename);
		if (shape == null)
		{
			slide.Delete();
			return null;
		}

		// center the picture
		shape.Top = (master.Height - shape.Height) / 2;
		shape.Left = (master.Width - shape.Width) / 2;

		return slide;
	}

	float slideAdvanceTime;
	public void AddPictureSlides(string path)
	{
		var files = Directory.GetFiles(path);

		int slides = files.Length;
		var pictureSlidesTime = mediaLegnth - (TitleAdvanceTime + TransitionDuration);
		var pictureSlideTime = (pictureSlidesTime / slides) - TransitionDuration;
		slideAdvanceTime = (float)pictureSlideTime.TotalSeconds;

		int count = 0; // During development, it can be convenient to stop after just a few pictures
		foreach (string file in files)
		{
			AddPictureSlide(file);

			count++;
			if (count > slides) break;
		}
	}

	public PPT.Slide AddEndSlide(string title, string subTitle)
	{
		var slide = AddSlide();

		SetText(slide.Shapes[1], title);
		SetText(slide.Shapes[2], subTitle);

		var transition = slide.SlideShowTransition;
		transition.AdvanceOnClick = Office.MsoTriState.msoTrue;
		transition.AdvanceOnTime = Office.MsoTriState.msoFalse;

		return slide;
	}

	public void CreateVideo(string fileName, int vertResolution = 720)
	{
		presentation.CreateVideo(fileName, UseTimingsAndNarrations: true, DefaultSlideDuration: 5, vertResolution, FramesPerSecond: 30, Quality: 85);

		// Yes, this is LAME ... but it's easy and works.
		while (true)
		{
			var status = presentation.CreateVideoStatus;
			if (status is PPT.PpMediaTaskStatus.ppMediaTaskStatusDone or PPT.PpMediaTaskStatus.ppMediaTaskStatusFailed)
			{
				break;
			}
			Thread.Sleep(1000);
		}
	}

	public void Close()
	{
		presentation.Close();
	}
}
