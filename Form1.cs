using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace xlsx2sql
{
    public partial class Form1 : Form
    {
     
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 点击导入数据库按钮，执行程序
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
           
           try
           {
               if ((textBox1.Text == null || textBox1.Text == "") && (textBox2.Text == null || textBox2.Text == ""))
            {
                label1.Text = "尚未选择文件或文件夹！";
            }
               else if (!(textBox1.Text == null || textBox1.Text == "") && (textBox2.Text == null || textBox2.Text == ""))

               { 
                   if (!textBox1.Text.Trim().EndsWith("xlsx") && !textBox1.Text.Trim().EndsWith("xls"))
                   {
                       label1.Text = "不支持的文件格式！";
                   }
                   else
                   {
                       //FileSvr fileSvr = new FileSvr();
                       //System.Data.DataTable dt = fileSvr.GetExcelDatatable("C:\\Users\\NewSpring\\Desktop\\Demo\\InExcelOutExcel\\InExcelOutExcel\\excel\\ExcelToDB.xlsx", "mapTable");
                       //System.Data.DataTable dt = fileSvr.GetExcelDatatable("D:\\ExcelToDB.xlsx", "mapTable");
                       //System.Data.DataTable dt = fileSvr.GetExcelDatatable(textBox1.Text, "Table");
                       //fileSvr.InsetData(dt);
                       //label1.Text = "操作完成！";
                       label1.Text = "";
                       startInset(textBox1.Text.Trim());
                       


                   }
               }
               else if ((textBox1.Text == null || textBox1.Text == "") && !(textBox2.Text == null || textBox2.Text == ""))
               {
                   string[] files = System.IO.Directory.GetFiles(textBox2.Text.Trim());
                   label1.Text = "";
                   foreach (string str in files)
                   {
                       startInset(str.Trim());
                       

                   }
                   label1.Text += "全部完成！";
               
               }
               
           }
           catch (Exception ex)
           {
               WriteLog.WriteError(ex.ToString());
               label1.Text = "出错！";
           }
        }


        /// <summary>
        /// 点击选择文件按钮，执行程序
        /// </summary>
         private void button2_Click(object sender, EventArgs e)
         {
             label1.Text = "";
             textBox1.Text = "";
             textBox2.Text = "";
            
             OpenFileDialog openFiledialog1 = new OpenFileDialog();
             
             openFiledialog1.ShowDialog();
             textBox1.Text = openFiledialog1.FileName;
             if (textBox1.Text == null || textBox1.Text == "")
             {
                 label1.Text = "尚未选择文件！";
             }
             else if (!textBox1.Text.Trim().EndsWith("xlsx") && !textBox1.Text.Trim().EndsWith("xls"))
             {
                 label1.Text = "不支持的文件格式！";
             }
             else
             {
                 
                 if (textBox1.Text.Trim().EndsWith("t_card.xlsx"))
                 {
                    
                     label1.Text = "已选择\"卡牌属性表\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_card_comb.xlsx"))
                 {
                     label1.Text = "已选择\"卡牌合成表\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_card_skill.xlsx"))
                 {
                     label1.Text = "已选择\"卡牌技能表\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_chapter_event.xlsx"))
                 {
                     label1.Text = "已选择\"t_chapter_event\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_element_gift.xlsx"))
                 {
                     label1.Text = "已选择\"t_element_gift\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_element_shop.xlsx"))
                 {
                     label1.Text = "已选择\"t_element_shop\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_event.xlsx"))
                 {
                     label1.Text = "已选择\"t_event\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_init_card_group.xlsx"))
                 {
                     label1.Text = "已选择\"t_init_card_group\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_item.xlsx"))
                 {
                     label1.Text = "已选择\"t_item\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_item_comb.xlsx"))
                 {
                     label1.Text = "已选择\"t_item_comb\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_item_enlarge.xlsx"))
                 {
                     label1.Text = "已选择\"t_item_enlarge\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_map.xlsx"))
                 {
                     label1.Text = "已选择\"t_map\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_map_group.xlsx"))
                 {
                     label1.Text = "已选择\"t_map_group\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_map_monster.xlsx"))
                 {
                     label1.Text = "已选择\"t_map_monster\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_monster.xlsx"))
                 {
                     label1.Text = "已选择\"t_monster\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_monster_card.xlsx"))
                 {
                     label1.Text = "已选择\"t_monster_card\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_monster_reward.xlsx"))
                 {
                     label1.Text = "已选择\"t_monster_reward\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_pvp_shop.xlsx"))
                 {
                     label1.Text = "已选择\"t_pvp_shop\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_reward_group.xlsx"))
                 {
                     label1.Text = "已选择\"t_reward_group\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_vip.xlsx"))
                 {
                     label1.Text = "已选择\"t_vip\"！";
                 }
                 else if (textBox1.Text.Trim().EndsWith("t_userLevel_conf.xlsx"))
                 {
                     label1.Text = "已选择\"t_userLevel_conf\"！";
                 }
                 else
                 {
                     label1.Text = "不支持的文件！";
                 }
             }
         }

         private void label1_Click(object sender, EventArgs e)
         {

         }
         /// <summary>
         /// 程序开始时运行
         /// </summary>
         private void Form1_Load(object sender, EventArgs e)
         {
             label1.Text = "请选择配置文件！";
             textBox1.ReadOnly = true;
             textBox2.ReadOnly = true;
             
         }
         /// <summary>
         /// 点击选择文件夹按钮，运行程序
         /// </summary>
         private void btnFolder_Click(object sender, EventArgs e)
         {
             label1.Text = "";
             textBox1.Text = "";
             textBox2.Text = "";
           
             FolderBrowserDialog fbd = new FolderBrowserDialog();
             fbd.Description = "请选择配置文件夹";
             if (fbd.ShowDialog() == DialogResult.OK)
             {
                 textBox2.Text = fbd.SelectedPath;
             }
            
         }
         /// <summary>
         /// 判断文件名，并调用对应的程序导入数据库
         /// </summary>
         private void startInset(string path)
         {
             try
             {
                 FileSvr fileSvr = new FileSvr();


                 if (path.Trim().EndsWith("t_card.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["card_id"].ToString().Trim() == dj["card_id"].ToString().Trim())
                             {
                                 error += di["card_id"].ToString() + "重复!\r\n";
                             }
                         }
                         if (di["card_id"].ToString() == null || di["card_id"].ToString() == "" || Regex.IsMatch(di["card_id"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_id" + "不符合规则!\r\n";
                         }
                         if (di["card_name"].ToString() == null || di["card_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "card_name" + "不符合规则!\r\n";
                         }
                         int t = di["card_id"].ToString().Last() - '1';
                         string[] quality = new string[] { "D", "C", "B", "A", "S", "S+", "SS", "SS+", "X" };
                         if (di["card_quality"].ToString() == null || di["card_quality"].ToString() == "" || di["card_quality"].ToString() != quality[t])
                         {
                             error += "第" + (i + 1) + "行" + "card_quality" + "不符合规则!\r\n";
                         }
                         if (di["card_HP"].ToString() == null || di["card_HP"].ToString() == "" || Regex.IsMatch(di["card_HP"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_HP" + "不符合规则!\r\n";
                         }
                         if (di["card_atk"].ToString() == null || di["card_atk"].ToString() == "" || Regex.IsMatch(di["card_atk"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_atk" + "不符合规则!\r\n";
                         }
                         if (di["card_level"].ToString() == null || di["card_level"].ToString() == "" || Regex.IsMatch(di["card_level"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_level" + "不符合规则!\r\n";
                         }
                         if (di["hp_growth"].ToString() == null || di["hp_growth"].ToString() == "" || Regex.IsMatch(di["hp_growth"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "hp_growth" + "不符合规则!\r\n";
                         }
                         if (di["atk_growth"].ToString() == null || di["atk_growth"].ToString() == "" || Regex.IsMatch(di["atk_growth"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "atk_growth" + "不符合规则!\r\n";
                         }
                         if (di["fight_growth"].ToString() == null || di["fight_growth"].ToString() == "" || Regex.IsMatch(di["fight_growth"].ToString().Trim(), @"^[0-9.]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "fight_growth" + "不符合规则!\r\n";
                         }
                         if (di["card_desc"].ToString() == null || di["card_desc"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "card_desc" + "不符合规则!\r\n";
                         }
                         if (di["card_race"].ToString() == null || di["card_race"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "card_race" + "不符合规则!\r\n";
                         }
                         if (di["card_attr"].ToString() == null || di["card_attr"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "card_attr" + "不符合规则!\r\n";
                         }
                         if (di["card_pic"].ToString() == null || di["card_pic"].ToString() == "" || Regex.IsMatch(di["card_pic"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false || di["card_pic"].ToString().Substring(0, 5) != di["card_id"].ToString().Substring(0, 5))
                         {
                             error += "第" + (i + 1) + "行" + "card_pic" + "不符合规则!\r\n";
                         }
                         if (di["head_pic"].ToString() == null || di["head_pic"].ToString() == "" || Regex.IsMatch(di["head_pic"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false || di["head_pic"].ToString().Substring(0, 5) != di["card_id"].ToString().Substring(0, 5))
                         {
                             error += "第" + (i + 1) + "行" + "head_pic" + "不符合规则!\r\n";
                         }
                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         int count1 = fileSvr.t_cardInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }


                 }
                 else if (path.Trim().EndsWith("t_card_comb.xlsx"))
                 {

                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["card_id"].ToString().Trim() == dj["card_id"].ToString().Trim())
                             {
                                 error += di["card_id"].ToString() + "重复!\r\n";
                             }
                         }
                         if (di["card_id"].ToString() == null || di["card_id"].ToString() == "" || Regex.IsMatch(di["card_id"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_id" + "不符合规则!\r\n";
                         }

                         if (di["material_id"].ToString() == null || di["material_id"].ToString() == "" || Regex.IsMatch(di["material_id"].ToString().Trim(), @"^[1][0-9]{5}$") == false || di["material_id"].ToString().Substring(1, 5) != di["card_id"].ToString().Substring(0, 5))
                         {
                             error += "第" + (i + 1) + "行" + "material_id" + "不符合规则!\r\n";
                         }
                         if (di["material_num"].ToString() == null || di["material_num"].ToString() == "" || Regex.IsMatch(di["material_num"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "material_num" + "不符合规则!\r\n";
                         }
                         if (di["need_gold"].ToString() == null || di["need_gold"].ToString() == "" || Regex.IsMatch(di["need_gold"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "need_gold" + "不符合规则!\r\n";
                         }
                         if (di["success_rate"].ToString() == null || di["success_rate"].ToString() == "" || Regex.IsMatch(di["success_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "success_rate" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {


                         int count1 = fileSvr.t_card_combInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_card_skill.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");

                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];

                         if (di["card_id"].ToString() == null || di["card_id"].ToString() == "" || Regex.IsMatch(di["card_id"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_id" + "不符合规则!\r\n";
                         }

                         if (di["skill_type"].ToString() == null || di["skill_type"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "skill_type" + "不符合规则!\r\n";
                         }
                         if (di["skill_name"].ToString() == null || di["skill_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "skill_name" + "不符合规则!\r\n";
                         }
                         if (di["description"].ToString() == null || di["description"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "description" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_card_skillInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_chapter_event.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");

                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];

                         if (di["chapter_id"].ToString() == null || di["chapter_id"].ToString() == "" || Regex.IsMatch(di["chapter_id"].ToString().Trim(), @"^[0-9]{3}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "chapter_id" + "不符合规则!\r\n";
                         }

                         if (di["event_id"].ToString() == null || di["event_id"].ToString() == "" || Regex.IsMatch(di["event_id"].ToString().Trim(), @"^[0-9]{8}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_id" + "不符合规则!\r\n";
                         }
                         if (di["event_rate"].ToString() == null || di["event_rate"].ToString() == "" || Regex.IsMatch(di["chapter_id"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_rate" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_chapter_eventInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_element_gift.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];

                         if (!(di["g_id"].ToString() == "1001" || di["g_id"].ToString() == "1002"))
                         {
                             error += "第" + (i + 1) + "行" + "g_id" + "不符合规则!\r\n";
                         }

                         if (di["g_type"].ToString() == null || di["g_type"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "g_type" + "不符合规则!\r\n";
                         }
                         if (di["g_value"].ToString() == null || di["g_value"].ToString() == "" || Regex.IsMatch(di["g_value"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "g_value" + "不符合规则!\r\n";
                         }
                         if (di["g_sign"].ToString() == null || di["g_sign"].ToString() == "" || Regex.IsMatch(di["g_sign"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "g_sign" + "不符合规则!\r\n";
                         }
                         if (di["g_num"].ToString() == null || di["g_num"].ToString() == "" || Regex.IsMatch(di["g_num"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "g_num" + "不符合规则!\r\n";
                         }
                         if (di["g_ratio"].ToString() == null || di["g_ratio"].ToString() == "" || Regex.IsMatch(di["g_ratio"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "g_ratio" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_element_giftInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_element_shop.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];

                         if (di["shop_id"].ToString() == null || di["shop_id"].ToString() == "" || Regex.IsMatch(di["shop_id"].ToString().Trim(), @"^[0-9]{6}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "shop_id" + "不符合规则!\r\n";
                         }

                         if (di["commodity"].ToString() == null || di["commodity"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "commodity" + "不符合规则!\r\n";
                         }
                         if (di["exp_type"].ToString() == null || di["exp_type"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "exp_type" + "不符合规则!\r\n";
                         }
                         if (di["exp_num"].ToString() == null || di["exp_num"].ToString() == "" || Regex.IsMatch(di["exp_num"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "exp_num" + "不符合规则!\r\n";
                         }
                         if (di["item_id"].ToString() == null || di["item_id"].ToString() == "" || Regex.IsMatch(di["item_id"].ToString().Trim(), @"^[0-9]{6}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "item_id" + "不符合规则!\r\n";
                         }
                         if (di["item_num"].ToString() == null || di["item_num"].ToString() == "" || Regex.IsMatch(di["item_num"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "item_num" + "不符合规则!\r\n";
                         }
                         if (di["free_time"].ToString() == null || di["free_time"].ToString() == "" || Regex.IsMatch(di["free_time"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "free_time" + "不符合规则!\r\n";
                         }
                         if (di["discount_pic"].ToString() == null || di["discount_pic"].ToString() == "" || Regex.IsMatch(di["discount_pic"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "discount_pic" + "不符合规则!\r\n";
                         }
                         if (!(di["g_id"].ToString() == "1001" || di["g_id"].ToString() == "1002"))
                         {
                             error += "第" + (i + 1) + "行" + "g_id" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_element_shopInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_event.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["event_id"].ToString().Trim() == dj["event_id"].ToString().Trim())
                             {
                                 error += di["event_id"].ToString() + "重复!\r\n";
                             }
                         }
                         if (di["event_id"].ToString() == null || di["event_id"].ToString() == "" || Regex.IsMatch(di["event_id"].ToString().Trim(), @"^[0-9]{8}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_id" + "不符合规则!\r\n";
                         }
                         if (di["event_name"].ToString() == null || di["event_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "event_name" + "不符合规则!\r\n";
                         }

                         if (di["event_desc"].ToString() == null || di["event_desc"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "event_desc" + "不符合规则!\r\n";
                         }
                         if (di["event_type"].ToString() == null || di["event_type"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "event_type" + "不符合规则!\r\n";
                         }
                         if (di["event_finish_type"].ToString() == null || di["event_finish_type"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "event_finish_type" + "不符合规则!\r\n";
                         }
                         if (!(di["event_isOpen"].ToString() == "0" || di["event_isOpen"].ToString() == "1"))
                         {
                             error += "第" + (i + 1) + "行" + "event_isOpen" + "不符合规则!\r\n";
                         }
                         if (di["event_canGiveup"].ToString() == null || di["event_canGiveup"].ToString() == "" || Regex.IsMatch(di["event_canGiveup"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_canGiveup" + "不符合规则!\r\n";
                         }

                         if (di["event_level"].ToString() == null || di["event_level"].ToString() == "" || Regex.IsMatch(di["event_level"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_level" + "不符合规则!\r\n";
                         }
                         if (di["event_rate"].ToString() == null || di["event_rate"].ToString() == "" || Regex.IsMatch(di["event_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_rate" + "不符合规则!\r\n";
                         }
                         if (di["event_pic"].ToString() == null || di["event_pic"].ToString() == "" || Regex.IsMatch(di["event_pic"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "event_pic" + "不符合规则!\r\n";
                         }
                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_eventInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_init_card_group.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];


                         if (di["group_id"].ToString() == null || di["group_id"].ToString() == "" || Regex.IsMatch(di["group_id"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "group_id" + "不符合规则!\r\n";
                         }
                         if (di["card_id"].ToString() == null || di["card_id"].ToString() == "" || Regex.IsMatch(di["card_id"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_id" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_init_card_groupInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_item.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["item_id"].ToString().Trim() == dj["item_id"].ToString().Trim())
                             {
                                 error += di["item_id"].ToString() + "重复!\r\n";
                             }
                         }
                         if (di["item_id"].ToString() == null || di["item_id"].ToString() == "" || Regex.IsMatch(di["item_id"].ToString().Trim(), @"^[0-9]{6}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "item_id" + "不符合规则!\r\n";
                         }
                         if (di["item_name"].ToString() == null || di["item_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "item_name" + "不符合规则!\r\n";
                         }
                         if (di["item_type"].ToString() == null || di["item_type"].ToString() == "" || Regex.IsMatch(di["item_type"].ToString().Trim(), @"^[1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "item_type" + "不符合规则!\r\n";
                         }
                         if (di["item_desc"].ToString() == null || di["item_desc"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "item_desc" + "不符合规则!\r\n";
                         }
                         if (di["type_desc"].ToString() == null || di["type_desc"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "type_desc" + "不符合规则!\r\n";
                         }

                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_itemInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_item_comb.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";



                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                       
                         int count1 =   fileSvr.t_item_combInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_item_enlarge.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["add_num"].ToString().Trim() == dj["add_num"].ToString().Trim())
                             {
                                 error += di["add_num"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["add_num"].ToString() == null || di["add_num"].ToString() == "" || Regex.IsMatch(di["add_num"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "add_num" + "不符合规则!\r\n";
                         }
                         if (di["need_gold"].ToString() == null || di["need_gold"].ToString() == "" || Regex.IsMatch(di["need_gold"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "need_gold" + "不符合规则!\r\n";
                         }
                         if (di["need_item"].ToString() == null || di["need_item"].ToString() == "" || Regex.IsMatch(di["need_item"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "need_item" + "不符合规则!\r\n";
                         }


                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_item_enlargeInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_map.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["map_id"].ToString().Trim() == dj["map_id"].ToString().Trim())
                             {
                                 error += di["map_id"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["map_id"].ToString() == null || di["map_id"].ToString() == "" || Regex.IsMatch(di["map_id"].ToString().Trim(), @"^[0-9]{5}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "map_id" + "不符合规则!\r\n";
                         }
                         if (di["map_name"].ToString() == null || di["map_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "map_name" + "不符合规则!\r\n";
                         }



                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_mapInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_map_group.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["map_group_id"].ToString().Trim() == dj["map_group_id"].ToString().Trim())
                             {
                                 error += di["map_group_id"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["map_group_id"].ToString() == null || di["map_group_id"].ToString() == "" || Regex.IsMatch(di["map_group_id"].ToString().Trim(), @"^[0-9]{3}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "map_group_id" + "不符合规则!\r\n";
                         }




                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_map_groupInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_map_monster.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];

                         if (di["map_id"].ToString() == null || di["map_id"].ToString() == "" || Regex.IsMatch(di["map_id"].ToString().Trim(), @"^[0-9]{5}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "map_id" + "不符合规则!\r\n";
                         }
                         if (di["monster_id"].ToString() == null || di["monster_id"].ToString() == "" || Regex.IsMatch(di["monster_id"].ToString().Trim(), @"^[0-9]{8}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "monster_id" + "不符合规则!\r\n";
                         }
                         if (di["mm_rate"].ToString() == null || di["mm_rate"].ToString() == "" || Regex.IsMatch(di["mm_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "mm_rate" + "不符合规则!\r\n";
                         }
                         if (di["reward_name"].ToString() == null || di["reward_name"].ToString() == "" || Regex.IsMatch(di["reward_name"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "reward_name" + "不符合规则!\r\n";
                         }




                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_map_monsterInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_monster.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["monster_id"].ToString().Trim() == dj["monster_id"].ToString().Trim())
                             {
                                 error += di["monster_id"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["monster_id"].ToString() == null || di["monster_id"].ToString() == "" || Regex.IsMatch(di["monster_id"].ToString().Trim(), @"^[0-9]{8}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "monster_id" + "不符合规则!\r\n";
                         }
                         if (di["monster_name"].ToString() == null || di["monster_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "monster_name" + "不符合规则!\r\n";
                         }
                         if (di["crit_rate"].ToString() == null || di["crit_rate"].ToString() == "" || Regex.IsMatch(di["crit_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "crit_rate" + "不符合规则!\r\n";
                         }
                         if (di["hit_rate"].ToString() == null || di["hit_rate"].ToString() == "" || Regex.IsMatch(di["hit_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "hit_rate" + "不符合规则!\r\n";
                         }
                         if (di["evade_rate"].ToString() == null || di["evade_rate"].ToString() == "" || Regex.IsMatch(di["evade_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "evade_rate" + "不符合规则!\r\n";
                         }
                         if (di["head_pic"].ToString() == null || di["head_pic"].ToString() == "" || Regex.IsMatch(di["head_pic"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "head_pic" + "不符合规则!\r\n";
                         }





                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_monsterInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_monster_card.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                       

                         if (di["monster_id"].ToString() == null || di["monster_id"].ToString() == "" || Regex.IsMatch(di["monster_id"].ToString().Trim(), @"^[0-9]{8}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "monster_id" + "不符合规则!\r\n";
                         }
                         if (di["card_id"].ToString() == null || di["card_id"].ToString() == "" || Regex.IsMatch(di["card_id"].ToString().Trim(), @"^[1-9][0-9]{4}[0][1-8]$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "card_id" + "不符合规则!\r\n";
                         }
                         if (di["card_name"].ToString() == null || di["card_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "card_name" + "不符合规则!\r\n";
                         }





                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_monster_cardInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_monster_reward.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];


                         if (di["rg_id"].ToString() == null || di["rg_id"].ToString() == "" || Regex.IsMatch(di["rg_id"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "rg_id" + "不符合规则!\r\n";
                         }
                         if (!(di["r_type"].ToString() == "A" || di["r_type"].ToString() == "B" || di["r_type"].ToString() == "C" || di["r_type"].ToString() == "D"))
                         {
                             error += "第" + (i + 1) + "行" + "r_type" + "不符合规则!\r\n";
                         }





                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_monster_rewardInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_pvp_shop.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["s_id"].ToString().Trim() == dj["s_id"].ToString().Trim())
                             {
                                 error += di["s_id"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["s_id"].ToString() == null || di["s_id"].ToString() == "" || Regex.IsMatch(di["s_id"].ToString().Trim(), @"^[0-9]{6}$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "s_id" + "不符合规则!\r\n";
                         }
                         if (di["s_name"].ToString() == null || di["s_name"].ToString() == "")
                         {
                             error += "第" + (i + 1) + "行" + "s_name" + "不符合规则!\r\n";
                         }





                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_pvp_shopInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_reward_group.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["id"].ToString().Trim() == dj["id"].ToString().Trim())
                             {
                                 error += di["id"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["id"].ToString() == null || di["id"].ToString() == "" || Regex.IsMatch(di["id"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "id" + "不符合规则!\r\n";
                         }
                         if (di["group_name"].ToString() == null || di["group_name"].ToString() == "" || Regex.IsMatch(di["group_name"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "s_name" + "不符合规则!\r\n";
                         }
                         if (di["reward_rate"].ToString() == null || di["reward_rate"].ToString() == "" || Regex.IsMatch(di["reward_rate"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "reward_rate" + "不符合规则!\r\n";
                         }
                         if (!(di["is_win"].ToString() == "1" || di["is_win"].ToString() == "0"))
                         {
                             error += "第" + (i + 1) + "行" + "id" + "不符合规则!\r\n";
                         }




                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_reward_groupInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_vip.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";

                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_vipInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else if (path.Trim().EndsWith("t_userLevel_conf.xlsx"))
                 {
                     System.Data.DataTable dt = fileSvr.GetExcelDatatable(path, "table");
                     //开始进行数据验证
                     int count = dt.Rows.Count;
                     string error = "共有" + count + "行数据！\r\n";
                     for (int i = 0; i < count; i++)
                     {
                         DataRow di = dt.Rows[i];
                         for (int j = i + 1; j < count; j++)
                         {

                             DataRow dj = dt.Rows[j];
                             if (di["levelNum"].ToString().Trim() == dj["levelNum"].ToString().Trim())
                             {
                                 error += di["levelNum"].ToString() + "重复!\r\n";
                             }
                         }

                         if (di["levelNum"].ToString() == null || di["levelNum"].ToString() == "" || Regex.IsMatch(di["levelNum"].ToString().Trim(), @"^[0-9]+$") == false)
                         {
                             error += "第" + (i + 1) + "行" + "levelNum" + "不符合规则!\r\n";
                         }
                        




                     }
                     //数据验证结束，提示结果，并选择是否继续导入
                     error += "是否继续导入数据？\r\n";
                     DialogResult result = MessageBox.Show(error, path, MessageBoxButtons.OKCancel, MessageBoxIcon.None);
                     //选择是，则开始导入数据
                     if (result == DialogResult.OK)
                     {
                         
                         int count1 = fileSvr.t_userLevel_confInsetData(dt);
                         if (count1 == count)
                         {
                             label1.Text += path + "操作完成！\r\n";
                         }
                         else
                         {
                             label1.Text += path + "插入数据条目不同，请检查！\r\n";
                         }
                     }
                 }
                 else
                 {
                     label1.Text += path + "不支持导入！\r\n";
                 }

             }
             catch (Exception ex)
             {
                 WriteLog.WriteError(ex.ToString());
                 label1.Text = "出错！";
             }
         }
        

         private void btnLog_Click(object sender, EventArgs e)
         {
             string 路径 = System.AppDomain.CurrentDomain.BaseDirectory.ToString().Trim() + "log";
             System.Diagnostics.Process.Start(路径);
         }
     
    }
 
}
