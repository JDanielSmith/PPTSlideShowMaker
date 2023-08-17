using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

internal sealed class Presentation
{
	readonly PPT.Presentation presentation;
	readonly PPT.Master master;
	public Presentation(PPT.Presentation presentation)
	{
		this.presentation = presentation;
		master = this.presentation.SlideMaster;
	}

	public TimeSpan TransitionDuration { get; init; } = TimeSpan.FromSeconds(0.85);

	public PPT.PpEntryEffect TransitionEntryEffect
	{
		set { Slides.TransitionEntryEffect = value; }
	}

	PPT.Slide AddSlide(int index = -1)
	{
		return Slides.AddSlide(presentation, TransitionDuration, index);
	}

	static void SetText(PPT.Shape shape, string text)
	{
		var textRange = shape.TextFrame.TextRange;
		textRange.Text = text;
		textRange.Font.Color.RGB = Int32.MaxValue; // White
	}

	public TimeSpan TitleAdvanceTime { get; init; } = TimeSpan.FromSeconds(4.0);

	public TimeSpan MediaLength { get; set; } = TimeSpan.FromMinutes(2.5); // in case there isn't any background music

	public PPT.Slide AddTitleSlide(string title, string? subTitle, string? backgroundMusicPathname)
	{
		var slide = AddSlide();
		slide.SlideShowTransition.AdvanceTime = (float)TitleAdvanceTime.TotalSeconds;

		SetText(slide.Shapes[1], title);
		if (subTitle is not null)
		{
			SetText(slide.Shapes[2], subTitle);
		}

		if (backgroundMusicPathname is not null)
		{
			MediaLength = Shapes.AddMediaObject2(slide.Shapes, Shortcut.Resolve(backgroundMusicPathname));
		}

		return slide;
	}

	public PPT.Slide? AddPictureSlide(string fileName, int index = -1)
	{
		var slide = AddSlide(index);
		Shapes.DeleteAll(slide.Shapes);
		slide.SlideShowTransition.AdvanceTime = SlideAdvanceTime;

		var shape = Shapes.AddPicture(slide.Shapes, Shortcut.Resolve(fileName));
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

	public float SlideAdvanceTime { get; set; }
	public PPT.Slide? AddEndSlide(string? title, string? subTitle)
	{
		if (title is null)
		{
			return null;
		}

		var slide = AddSlide();
		SetText(slide.Shapes[1], title);
		if (subTitle is not null)
		{
			SetText(slide.Shapes[2], subTitle);
		}

		var transition = slide.SlideShowTransition;
		transition.AdvanceOnClick = Office.MsoTriState.msoTrue;
		transition.AdvanceOnTime = Office.MsoTriState.msoFalse;

		return slide;
	}

	PPT.PpMediaTaskStatus GetCreateVideoStatus()
	{
		try
		{
			return presentation.CreateVideoStatus;
		}
		catch (COMException)
		{
			return PPT.PpMediaTaskStatus.ppMediaTaskStatusNone;
		}
	}

	public void CreateVideo(string fileName, int vertResolution = 720)
	{
		// Don't need a high frame-rate as these are still photos.
		presentation.CreateVideo(fileName, UseTimingsAndNarrations: true, DefaultSlideDuration: 5, vertResolution, FramesPerSecond: 7, Quality: 80);

		// Yes, this is LAME ... but it's easy and "works."
		while (true)
		{
			var status = GetCreateVideoStatus();
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
