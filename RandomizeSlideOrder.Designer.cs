namespace RandomizeSlideOrder
{
    partial class RandomizeSlideOrder : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RandomizeSlideOrder()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RandomizeSlideOrder));
            this.tabSlideOrder = this.Factory.CreateRibbonTab();
            this.slideOrder = this.Factory.CreateRibbonGroup();
            this.shuffleAndStart = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.shuffleAndLoop = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.resetButton = this.Factory.CreateRibbonButton();
            this.tabSlideOrder.SuspendLayout();
            this.slideOrder.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSlideOrder
            // 
            this.tabSlideOrder.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSlideOrder.ControlId.OfficeId = "TabSlideShow";
            this.tabSlideOrder.Groups.Add(this.slideOrder);
            this.tabSlideOrder.Label = "TabSlideShow";
            this.tabSlideOrder.Name = "tabSlideOrder";
            // 
            // slideOrder
            // 
            this.slideOrder.Items.Add(this.shuffleAndStart);
            this.slideOrder.Items.Add(this.separator1);
            this.slideOrder.Items.Add(this.shuffleAndLoop);
            this.slideOrder.Items.Add(this.separator2);
            this.slideOrder.Items.Add(this.resetButton);
            this.slideOrder.Label = "Slide Order";
            this.slideOrder.Name = "slideOrder";
            this.slideOrder.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupSlideShowStart");
            // 
            // shuffleAndStart
            // 
            this.shuffleAndStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.shuffleAndStart.Image = global::RandomizeSlideOrder.Properties.Resources.shuffle;
            this.shuffleAndStart.Label = "Shuffle Start";
            this.shuffleAndStart.Name = "shuffleAndStart";
            this.shuffleAndStart.ScreenTip = "Shuffles the order of the presentation and starts it";
            this.shuffleAndStart.ShowImage = true;
            this.shuffleAndStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShuffleOrderAndPlay);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // shuffleAndLoop
            // 
            this.shuffleAndLoop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.shuffleAndLoop.Image = global::RandomizeSlideOrder.Properties.Resources.loop;
            this.shuffleAndLoop.Label = "Shuffle Loop";
            this.shuffleAndLoop.Name = "shuffleAndLoop";
            this.shuffleAndLoop.ScreenTip = "Shuffles te presentation order and continuously loops until exiting";
            this.shuffleAndLoop.ShowImage = true;
            this.shuffleAndLoop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoopOrder);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // resetButton
            // 
            this.resetButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.resetButton.Image = ((System.Drawing.Image)(resources.GetObject("resetButton.Image")));
            this.resetButton.Label = "Reset Order";
            this.resetButton.Name = "resetButton";
            this.resetButton.ShowImage = true;
            this.resetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ResetOrder);
            // 
            // RandomizeSlideOrder
            // 
            this.Name = "RandomizeSlideOrder";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabSlideOrder);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LoadRibbon);
            this.tabSlideOrder.ResumeLayout(false);
            this.tabSlideOrder.PerformLayout();
            this.slideOrder.ResumeLayout(false);
            this.slideOrder.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSlideOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup slideOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton shuffleAndStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton shuffleAndLoop;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
    }

    partial class ThisRibbonCollection
    {
        internal RandomizeSlideOrder Ribbon1
        {
            get { return this.GetRibbon<RandomizeSlideOrder>(); }
        }
    }
}
