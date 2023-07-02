namespace ToJson_Winform
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_FolderBrowser = new System.Windows.Forms.Button();
            this.btn_TableFileSelectBrowser = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_FolderBrowser
            // 
            this.btn_FolderBrowser.Location = new System.Drawing.Point(53, 56);
            this.btn_FolderBrowser.Name = "btn_FolderBrowser";
            this.btn_FolderBrowser.Size = new System.Drawing.Size(210, 64);
            this.btn_FolderBrowser.TabIndex = 0;
            this.btn_FolderBrowser.Text = "테이블 폴더 선택";
            this.btn_FolderBrowser.UseVisualStyleBackColor = true;
            this.btn_FolderBrowser.Click += new System.EventHandler(this.folderAndFileBrowsing);
            // 
            // btn_TableFileSelectBrowser
            // 
            this.btn_TableFileSelectBrowser.Location = new System.Drawing.Point(287, 56);
            this.btn_TableFileSelectBrowser.Name = "btn_TableFileSelectBrowser";
            this.btn_TableFileSelectBrowser.Size = new System.Drawing.Size(218, 64);
            this.btn_TableFileSelectBrowser.TabIndex = 1;
            this.btn_TableFileSelectBrowser.Text = "테이블 파일 선택";
            this.btn_TableFileSelectBrowser.UseVisualStyleBackColor = true;
            this.btn_TableFileSelectBrowser.Click += new System.EventHandler(this.tableFileBrowsing);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(18F, 30F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1521, 1492);
            this.Controls.Add(this.btn_TableFileSelectBrowser);
            this.Controls.Add(this.btn_FolderBrowser);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_FolderBrowser;
        private System.Windows.Forms.Button btn_TableFileSelectBrowser;
    }
}

