internal static class Drawing
{
	/// <summary>
	/// Does this given filename appear to be an image?  That is, can
	/// `System.Drawing.Image.FromFile()` load it?
	/// </summary>
	/// <param name="filename">Image filename to test.</param>
	/// <returns>Whether or not the filename was loaded as an image.</returns>
	public static bool IsImage(string filename)
	{
		try
		{
			using var image = System.Drawing.Image.FromFile(filename);
			return true;
		}
		catch (OutOfMemoryException)
		{
			return false;
		}
	}

}
