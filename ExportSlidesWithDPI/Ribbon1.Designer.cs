namespace ExportSlidesWithDPIDoing
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.Button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.comboBox2 = this.Factory.CreateRibbonComboBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.label4 = this.Factory.CreateRibbonLabel();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.label5 = this.Factory.CreateRibbonLabel();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Label = "Pic Export";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.Button1);
            this.group1.Items.Add(this.button2);
            this.group1.Name = "group1";
            // 
            // label1
            // 
            this.label1.Label = "Save";
            this.label1.Name = "label1";
            // 
            // Button1
            // 
            this.Button1.Description = "请选择图片输出位置";
            this.Button1.Label = "Path";
            this.Button1.Name = "Button1";
            this.Button1.ScreenTip = "Please select the image output location";
            // 
            // button2
            // 
            this.button2.Label = "Export";
            this.button2.Name = "button2";
            this.button2.SuperTip = "Performing an export operation";
            // 
            // group2
            // 
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.comboBox1);
            this.group2.Name = "group2";
            // 
            // label2
            // 
            this.label2.Label = "DPI";
            this.label2.Name = "label2";
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "DPI";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.SuperTip = "Set the DPI of the selected image";
            this.comboBox1.Text = null;
            // 
            // group3
            // 
            this.group3.Items.Add(this.label3);
            this.group3.Items.Add(this.comboBox2);
            this.group3.Name = "group3";
            // 
            // label3
            // 
            this.label3.Label = "Format";
            this.label3.Name = "label3";
            // 
            // comboBox2
            // 
            this.comboBox2.Label = "Format";
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.SuperTip = "Set the format of the selected image";
            this.comboBox2.Text = null;
            // 
            // group4
            // 
            this.group4.Items.Add(this.label4);
            this.group4.Items.Add(this.editBox1);
            this.group4.Name = "group4";
            // 
            // label4
            // 
            this.label4.Label = "Page";
            this.label4.Name = "label4";
            // 
            // editBox1
            // 
            this.editBox1.Label = "Page";
            this.editBox1.Name = "editBox1";
            this.editBox1.SuperTip = "Page number (format example: all or 0 or 1-3,5)";
            this.editBox1.Text = null;
            // 
            // group5
            // 
            this.group5.Items.Add(this.label5);
            this.group5.Items.Add(this.button3);
            this.group5.Items.Add(this.button4);
            this.group5.Name = "group5";
            // 
            // label5
            // 
            this.label5.Label = "About";
            this.label5.Name = "label5";
            // 
            // button3
            // 
            this.button3.Label = "Dev";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Label = "About";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click_1);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label4;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
