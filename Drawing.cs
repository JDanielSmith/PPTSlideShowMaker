using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

internal static class Drawing
{
	public static bool IsImage(string filename)
	{
		try
		{
			using var image = Image.FromFile(filename);
		}
		catch (OutOfMemoryException)
		{
			return false;
		}
		return true;
	}

}
