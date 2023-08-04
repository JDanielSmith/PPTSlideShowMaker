using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

internal static class Shortcut
{

	static string GetTargetPath(string lnkPath)
	{
		var shell = new IWshRuntimeLibrary.WshShell();
		dynamic shortcut = shell.CreateShortcut(lnkPath);
		return shortcut.TargetPath;
	}
	static string Resolve_(string path)
	{
		var extension = Path.GetExtension(path);
		if (String.Equals(extension, ".lnk", StringComparison.InvariantCultureIgnoreCase))
		{
			return GetTargetPath(path);
		}
		return path;
	}
	public static string Resolve(string path)
	{
		var retval = Resolve_(path);
		if (!File.Exists(retval))
		{
			throw new FileNotFoundException(retval);
		}
		return retval;
	}

}
