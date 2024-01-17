using System.Diagnostics;

internal static class TagEditor
{
	static string? FindTagEditor(DirectoryInfo? path)
	{
		if (path is null)
		{
			return null;
		}

		const string tagEditorDirectoryName = @"tageditor-3.5.0";
		const string tagEditorExeName = @"tageditor-3.5.0-x86_64-w64-mingw32.exe";

		string retval = Path.Combine(path.FullName, tagEditorDirectoryName, tagEditorExeName);
		if (File.Exists(retval))
		{
			return retval;
		}

		return FindTagEditor(path.Parent);
	}

	static Process Launch_(string exe, string file, string cover)
	{
		// https://www.addictivetips.com/windows-tips/set-thumbnail-image-for-a-video-on-windows-10/
		// -s cover=my-cover.jpg --max-padding 125000 -f My_video.mp4
		// https://github.com/Martchus/tageditor
		/*
		 tageditor set title="Title of "{1st,2nd,3rd}" file" title="Title of "{4..16}"th file" \
			  album="The Album" artist="The Artist" \
			  cover'=/path/to/image' lyrics'>=/path/to/lyrics' track'+=1/16' --files /some/dir/*.m4a
		*/
		var args = new string[] {
				"set", "cover=" +cover,
				"--max-padding", "125000",
				 "-f", file,
				};
		//var args = new string[] {"--help" };
		return Process.Start(exe, args);
	}
	static Process Launch(string exe, string path, FileInfo cover)
	{
		var currentDirectory = Environment.CurrentDirectory;
		Environment.CurrentDirectory = cover.DirectoryName!;
		try
		{
			return Launch_(exe, path, cover.Name);
		}
		finally
		{
			Environment.CurrentDirectory = currentDirectory;
		}
	}

	public static void AddCover(string path, FileInfo cover)
	{
		var tagEditorExePath = FindTagEditor(new DirectoryInfo(Environment.CurrentDirectory));
		if (tagEditorExePath is null)
		{
			return;
		}

		using (var tagEditor = Launch(tagEditorExePath, path, cover))
		{
			tagEditor.WaitForExit();
		}
	}
}

