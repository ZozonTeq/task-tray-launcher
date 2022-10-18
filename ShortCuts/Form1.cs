using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShortCuts
{

    public partial class Form1 : Form
    {
        public static class Texts//テキスト格納
        {
            public static string Replace = "置き換え";
            public static string Exit = "終了";
            public static string Restart = "再起動";
            public static string Reload = "再読み込み";
            public static string Open = "開く";
        }
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)//読み込み時
        {
            Hide();
            notifyIcon1.Visible = true;
            if (!Directory.Exists("shortcuts")) Directory.CreateDirectory("shortcuts");//ショートカットをぶち込むディレクトリ作成。
            if (File.Exists("logo.ico"))// カスタムアイコンの設定。
            {
                string iconPath = "logo.ico";
                Icon icon = new Icon(iconPath,48,48);
                notifyIcon1.Icon = icon;
            }
        }
        private void contextMenuStrip1_Closed(object sender, ToolStripDropDownClosedEventArgs e)
        {
            contextMenuStrip1 = new ContextMenuStrip();
            contextMenuStrip1.Opened += contextMenuStrip1_Opened;
            contextMenuStrip1.Items.Add("Exit").Name="Exit";
            contextMenuStrip1.Items.Add("Reload").Name = "Reload";
            contextMenuStrip1.Items.Add("Open").Name = "Open";
        }
        private void initContextMenu()//コンテクストメニューの初期化
        {
            contextMenuStrip1.Items.Clear();
            string[] files = Directory.GetFiles("shortcuts");
            for (int i = 0; i < files.Length; i++)
            {
                string ext = Path.GetExtension(files[i].ToString());//拡張子

                contextMenuStrip1.Items.Add(files[i]);//コンテクストメニューにアイテム追加
                contextMenuStrip1.Items[i].Name = files[i];//名前を設定
                contextMenuStrip1.Items[i].Text = files[i].Replace("shortcuts\\", "");//表示名の設定
                contextMenuStrip1.Items[i].Click += OnPressFile;//イベントの設定


                if (ext == ".lnk")//アイテムがショートカットだった場合
                {
                    string path = "";
                    bool success = true;
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    try
                    {
                        IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(files[i].ToString());
                        string targetPath = shortcut.TargetPath.ToString();
                        path = files[i];
                        Icon appIcon = Icon.ExtractAssociatedIcon(shortcut.TargetPath);
                        contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                    }//ショートカット先のパスを取得しビットマップを取得
                    catch (Exception e)
                    {
                        success = false;
                    }
                    if (!success)
                    {
                        try
                        {
                            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(files[i].ToString());
                            string targetPath = shortcut.TargetPath.ToString();

                            path = files[i];
                            Icon appIcon = Icon.ExtractAssociatedIcon(shortcut.TargetPath.Replace("C:\\Program Files (x86)\\", "C:\\Program Files\\"));
                            contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();//イメージを設定
                        }//(x86)を消してビットマップ取得。
                        catch (Exception e)
                        {
                            success = false;
                        }
                    }
                    else if (!success)
                    {
                        try
                        {
                            path = files[i];
                            Icon appIcon = Icon.ExtractAssociatedIcon(path);
                            contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();//イメージを設定
                            success = true;
                        }
                        catch (Exception e)
                        {
                            success = false;
                        }
                    }
                }//ショートカットファイルだったら
                else
                {
                    string path = files[i];
                    Icon appIcon = System.Drawing.Icon.ExtractAssociatedIcon(path);
                    contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                }
                //画像の設定終わり

            }
            contextMenuStrip1.Items.Add("Splitter").Name = "Splitter";
            contextMenuStrip1.Items.Find("Splitter", false)[0].Text = " ";

            contextMenuStrip1.Items.Add("Open").Name = "Open";
            contextMenuStrip1.Items.Find("Open", false)[0].Text = Texts.Open;
            contextMenuStrip1.Items.Find("Open", false)[0].Click += OnPressOpen;

            contextMenuStrip1.Items.Add("Reload").Name = "Reload";
            contextMenuStrip1.Items.Find("Reload", false)[0].Text = Texts.Reload;
            contextMenuStrip1.Items.Find("Reload", false)[0].Click += OnPressReload;

            contextMenuStrip1.Items.Add("Exit").Name = "Exit";
            contextMenuStrip1.Items.Find("Exit", false)[0].Text = Texts.Exit;
            contextMenuStrip1.Items.Find("Exit", false)[0].Click += OnPressExit;

            for (int i = 0; i < contextMenuStrip1.Items.Count; i++)
            {
                contextMenuStrip1.Items[i].ForeColor = Color.Black;
                contextMenuStrip1.Items[i].BackColor = Color.White;
                contextMenuStrip1.Items[i].Margin = Padding.Empty;
            }//アイテムの見た目設定。
        }

        private void contextMenuStrip1_Opened(object sender, EventArgs e)
        {
            initContextMenu();//コンテクストメニューの初期化
        }
        private void OnPressExit(object sender,EventArgs e)=> Close();//プロセス終了

        private void OnPressFile(object sender,EventArgs e)//ショートカットアイテムクリック時
        {
            string tag = "shortcuts\\"+sender.ToString();
            Console.WriteLine(tag);
            string ext = Path.GetExtension(tag);//拡張子
            if(ext == ".lnk")
            {
                IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(tag);
                string targetPath = shortcut.TargetPath.ToString();
                bool success = true;
                try
                {
                    System.Diagnostics.Process.Start(targetPath);//普通に実行
                }
                catch (Exception ex)
                {
                    success = false;
                }
                if (!success)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(targetPath.Replace("C:\\Program Files (x86)\\", "C:\\Program Files\\"));
                    }//Program Files (x86)をProgram Filesに置き換えて実行。
                    catch (Exception ex)
                    {
                        success=false;
                    }
                }
                if (!success)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(tag);
                    }
                    catch (Exception ex)
                    {
                        success = false;
                    }
                }
            }
            else//.lnk以外の場合
            {
                try
                {
                    System.Diagnostics.Process.Start(tag);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }//通常の実行
        }
        private void OnPressOpen(object sender, EventArgs e)//開くボタンを押したとき
        {
            System.Diagnostics.Process.Start("EXPLORER.EXE", $@"shortcuts");
            //Console.WriteLine($"{Directory.GetCurrentDirectory() + "\\shortcuts\\"}");
        }
        private void OnPressReload(object sender,EventArgs e)//再読み込みボタン押したとき
        {
            Application.Restart();//再起動する。
        }


    }
}
