using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

internal static class Shapes
{
	public static void DeleteAll(PPT.Shapes shapes)
	{
		while (shapes.Count > 0)
		{
			shapes[1].Delete();
		}
	}

	public static TimeSpan AddMediaObject2(PPT.Shapes shapes, string fileName)
	{
		var linkToFile = Office.MsoTriState.msoTrue;
		var saveWithDocument = Office.MsoTriState.msoFalse;
		var shape = shapes.AddMediaObject2(Shortcut.Resolve(fileName), linkToFile, saveWithDocument);

		var mediaLength = TimeSpan.FromMilliseconds(shape.MediaFormat.Length);

		// https://learn.microsoft.com/en-us/office/vba/api/powerpoint.playsettings
		var animationSettings = shape.AnimationSettings;
		animationSettings.AdvanceMode = PPT.PpAdvanceMode.ppAdvanceOnTime;
		var playSettings = animationSettings.PlaySettings;
		playSettings.PlayOnEntry = Office.MsoTriState.msoTrue;
		playSettings.PauseAnimation = Office.MsoTriState.msoFalse;
		playSettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
		playSettings.StopAfterSlides = Int32.MaxValue;

		return mediaLength;
	}

	public static PPT.Shape? AddPicture(PPT.Shapes shapes, string fileName)
	{
		// PPT will take just about anything as an "image," try to ensure
		// we only add actual images.
		if (!Drawing.IsImage(fileName))
		{
			return null;
		}

		var linkToFile = Office.MsoTriState.msoTrue;
		var saveWithDocument = Office.MsoTriState.msoFalse;
		try
		{
			return shapes.AddPicture(Shortcut.Resolve(fileName), linkToFile, saveWithDocument, Left: 0, Top: 0);
		}
		catch (System.Runtime.InteropServices.COMException)
		{
			return null;
		}
	}

}

