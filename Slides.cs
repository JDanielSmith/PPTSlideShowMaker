using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

internal static class Slides
{

	static PPT.CustomLayout GetLayout(PPT.Presentation presentation)
	{
		var master = presentation.SlideMaster;
		return master.CustomLayouts[1];
	}

	public static PPT.PpEntryEffect TransitionEntryEffect { get; set; } = PPT.PpEntryEffect.ppEffectRandom;
	public static PPT.Slide AddSlide(PPT.Presentation presentation, TimeSpan transitionDuration, int index = -1)
	{
		if (index < 0)
		{
			index = presentation.Slides.Count + 1;
		}
		var slide = presentation.Slides.AddSlide(index, GetLayout(presentation));
		slide.ColorScheme[PPT.PpColorSchemeIndex.ppBackground].RGB = Int32.MinValue; // Black

		var transition = slide.SlideShowTransition;
		transition.EntryEffect = TransitionEntryEffect;
		transition.Duration = (float)transitionDuration.TotalSeconds;
		transition.AdvanceOnClick = Office.MsoTriState.msoFalse;
		transition.AdvanceOnTime = Office.MsoTriState.msoTrue;

		return slide;
	}
}

