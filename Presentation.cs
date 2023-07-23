using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PPTSlideshowMaker
{
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

        public TimeSpan TransitionDuration { get; set; } = TimeSpan.FromSeconds(1);

        static void DeleteAllShapes(PPT.Slide slide)
        {
            while (slide.Shapes.Count > 0)
            {
                slide.Shapes[1].Delete();
            }
        }

        PPT.Slide AddSlide(int index = -1)
        {
            if (index < 0)
            {
                index = presentation.Slides.Count + 1;
            }
            var slide = presentation.Slides.AddSlide(index, layout);
            slide.ColorScheme[PPT.PpColorSchemeIndex.ppBackground].RGB = Int32.MinValue; // Black

            var transition = slide.SlideShowTransition;
            transition.EntryEffect = PPT.PpEntryEffect.ppEffectRandom;
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

        public TimeSpan TitleAdvanceTime { get; set; } = TimeSpan.FromSeconds(3.0);

        TimeSpan MediaLegnth { get; set; } = TimeSpan.Zero;
        public PPT.Slide AddTitleSlide(string title, string subTitle, string backgroundMusicPathname)
        {
            var slide = AddSlide();
            slide.SlideShowTransition.AdvanceTime = (float)TitleAdvanceTime.TotalSeconds;

            SetText(slide.Shapes[1], title);
            SetText(slide.Shapes[2], subTitle);

            var linkToFile = Office.MsoTriState.msoTrue;
            var saveWithDocument = Office.MsoTriState.msoFalse;
            var shape = slide.Shapes.AddMediaObject2(backgroundMusicPathname, linkToFile, saveWithDocument);

            MediaLegnth = TimeSpan.FromMilliseconds(shape.MediaFormat.Length);

            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.playsettings
            var animationSettings = shape.AnimationSettings;
            animationSettings.AdvanceMode = PPT.PpAdvanceMode.ppAdvanceOnTime;
            var playSettings = animationSettings.PlaySettings;
            playSettings.PlayOnEntry = Office.MsoTriState.msoTrue;
            playSettings.LoopUntilStopped = Office.MsoTriState.msoTrue;
            playSettings.PauseAnimation = Office.MsoTriState.msoFalse;
            playSettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
            playSettings.StopAfterSlides = 9999;

            return slide;
        }

        public PPT.Slide AddPictureSlide(string filename, int index = -1)
        {
            var slide = AddSlide();
            DeleteAllShapes(slide);
            slide.SlideShowTransition.AdvanceTime = slideAdvanceTime;

            var linkToFile = Office.MsoTriState.msoTrue;
            var saveWithDocument = Office.MsoTriState.msoFalse;
            var shape = slide.Shapes.AddPicture(filename, linkToFile, saveWithDocument, Left: 0, Top: 0);

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
            var pictureSlidesTime = MediaLegnth - (TitleAdvanceTime + TransitionDuration);
            var pictureSlideTime = (pictureSlidesTime / slides) - TransitionDuration;
            slideAdvanceTime = (float)pictureSlideTime.TotalSeconds;

            int count = 0;
            foreach (string file in files)
            {
                AddPictureSlide(file);

                count++;
                if (count > slides) break;
            }
        }

        public void CreateVideo(string filename, int vertResolution = 720)
        {
            presentation.CreateVideo(filename, UseTimingsAndNarrations: true, DefaultSlideDuration: 5, vertResolution, FramesPerSecond: 30, Quality: 85);

            while (true)
            {
                var status = presentation.CreateVideoStatus;
                if (status != PPT.PpMediaTaskStatus.ppMediaTaskStatusDone)
                {
                    Thread.Sleep(5000);
                }
                else
                {
                    break;
                }
            }
        }

        public void Close()
        {
            presentation.Close();
        }
    }
}
