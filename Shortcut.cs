internal static class Shortcut
{
	static string GetTargetPath(string lnkPath)
	{
		var shell = new IWshRuntimeLibrary.WshShell();
		dynamic shortcut = shell.CreateShortcut(lnkPath);
		return shortcut.TargetPath;
	}

	public static bool IsShortcut(string path)
	{
		var extension = Path.GetExtension(path);
		return String.Equals(extension, ".lnk", StringComparison.InvariantCultureIgnoreCase);
	}
	public static string Resolve(string path)
	{
		var retval = IsShortcut(path) ? GetTargetPath(path) : path;
		if (!File.Exists(retval))
		{
			throw new FileNotFoundException(retval);
		}
		return retval;
	}

}
