using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace RandomizeSlideOrder
{
    public partial class RandomizeSlideOrder
    {
        Application powerpoint;
        private void LoadRibbon(object sender, RibbonUIEventArgs e)
        {
            
        }
        private void ShuffleOrderAndPlay(object sender, RibbonControlEventArgs e)
        {
            powerpoint = Globals.ThisAddIn.Application;
            powerpoint.StartNewUndoEntry();
            SlideShowSettings presentation = ShuffleOrder();
            if (presentation == null) return;
            presentation.Run();
        }
        private void LoopOrder(object sender, RibbonControlEventArgs e)
        {
            powerpoint = Globals.ThisAddIn.Application;
            powerpoint.StartNewUndoEntry();
            SlideShowSettings presentation = LoopOrder();
            if (presentation == null) return;
            presentation.Run();
            
        }
        private void ResetOrder(object sender, RibbonControlEventArgs e)
        {
            if(powerpoint == null) return;
            try
            {
                powerpoint.CommandBars.ExecuteMso("Undo");
            }
            catch(Exception) {}
        }
        private SlideShowSettings ShuffleOrder()
        {
            powerpoint = Globals.ThisAddIn.Application;
            Slides slides = powerpoint.ActivePresentation.Slides;
            if (slides.Count > 0)
            {
                Random random = new Random();
                for (int i = 1; i <= slides.Count; i++)
                {
                    try
                    {
                        int randNum = random.Next(i, slides.Count);
                        slides[randNum].Cut();
                        slides.Paste(1);
                    }
                    catch (Exception) {}
                }
                return powerpoint.ActivePresentation.SlideShowSettings;
            }
            return null;
        }
        private SlideShowSettings LoopOrder()
        {
            powerpoint = Globals.ThisAddIn.Application;
            powerpoint.StartNewUndoEntry();
            Slides slides = powerpoint.ActivePresentation.Slides;
            if (slides.Count > 0)
            {
                Random random = new Random();
                for (int i = 1; i <= slides.Count; i++)
                {
                    int randNum = random.Next(i, slides.Count);
                    slides[randNum].Cut();
                    slides.Paste(1);
                }
                powerpoint.ActivePresentation.SlideShowSettings.LoopUntilStopped = MsoTriState.msoTrue;
                return powerpoint.ActivePresentation.SlideShowSettings;
            }
            return null;
        }
    }
}
