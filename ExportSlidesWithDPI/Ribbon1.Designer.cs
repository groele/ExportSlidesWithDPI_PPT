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
            this.PicExport = this.Factory.CreateRibbonTab();
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
            this.group6 = this.Factory.CreateRibbonGroup();
            this.label6 = this.Factory.CreateRibbonLabel();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.label5 = this.Factory.CreateRibbonLabel();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.PicExport.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // PicExport
            // 
            this.PicExport.Groups.Add(this.group1);
            this.PicExport.Groups.Add(this.group2);
            this.PicExport.Groups.Add(this.group3);
            this.PicExport.Groups.Add(this.group4);
            this.PicExport.Groups.Add(this.group6);
            this.PicExport.Groups.Add(this.group5);
            this.PicExport.Label = "输出图片";
            this.PicExport.Name = "PicExport";
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
            this.label1.Label = "文件";
            this.label1.Name = "label1";
            // 
            // Button1
            // 
            this.Button1.Description = "请选择图片输出位置";
            this.Button1.Label = "文件夹";
            this.Button1.Name = "Button1";
            this.Button1.ScreenTip = "请选择图片存储位置";
            // 
            // button2
            // 
            this.button2.Label = "导出";
            this.button2.Name = "button2";
            this.button2.SuperTip = "执行导出操作";
            // 
            // group2
            // 
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.comboBox1);
            this.group2.Name = "group2";
            // 
            // label2
            // 
            this.label2.Label = "图片分辨率";
            this.label2.Name = "label2";
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "分辨率";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.SuperTip = "建议300dpi，600dpi第一次到处会有30s左右卡顿，后续速度提升";
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
            this.label3.Label = "图片格式";
            this.label3.Name = "label3";
            // 
            // comboBox2
            // 
            this.comboBox2.Label = "格式";
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.SuperTip = "请选择导出图片的格式";
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
            this.label4.Label = "导出范围";
            this.label4.Name = "label4";
            // 
            // editBox1
            // 
            this.editBox1.Label = "页码";
            this.editBox1.Name = "editBox1";
            this.editBox1.SuperTip = "格式示例：all：导出所有图片；0：导出当前图片；1-3,5：导出所选择图片范围";
            this.editBox1.Text = null;
            // 
            // group6
            // 
            this.group6.Items.Add(this.label6);
            this.group6.Items.Add(this.checkBox1);
            this.group6.Items.Add(this.editBox2);
            this.group6.Name = "group6";
            // 
            // label6
            // 
            this.label6.Label = "裁剪白边";
            this.label6.Name = "label6";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "启用裁剪";
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.SuperTip = "启用自动裁剪图片四周的白边";
            // 
            // editBox2
            // 
            this.editBox2.Label = "留白大小";
            this.editBox2.Name = "editBox2";
            this.editBox2.SuperTip = "设置裁剪后保留的边距大小（单位：像素）\n建议值：0-50像素";
            this.editBox2.Text = "0";
            // 
            // group5
            // 
            this.group5.Items.Add(this.label5);
            this.group5.Items.Add(this.button4);
            this.group5.Items.Add(this.button3);
            this.group5.Name = "group5";
            // 
            // label5
            // 
            this.label5.Label = "帮助";
            this.label5.Name = "label5";
            // 
            // button4
            // 
            this.button4.Label = "使用说明";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click_1);
            // 
            // button3
            // 
            this.button3.Label = "关于";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.PicExport);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.PicExport.ResumeLayout(false);
            this.PicExport.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PicExport;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label6;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
