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

namespace xlsx2sql
{

    class FileSvr
    {
            /// <summary>
            /// Excel数据导入Datable
            /// </summary>
            /// <param name="fileUrl"></param>
            /// <param name="table"></param>
            /// <returns></returns>
            public System.Data.DataTable GetExcelDatatable(string fileUrl, string table)
            {
                //office2007之前 仅支持.xls
                //const string cmdText = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;IMEX=1';";
                //支持.xls和.xlsx，即包括office2010等版本的   HDR=Yes代表第一行是标题，不是数据；
                const string cmdText = "Provider=Microsoft.Ace.OleDb.12.0;Data Source={0};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
                System.Data.DataTable dt = null;
                //建立连接
                OleDbConnection conn = new OleDbConnection(string.Format(cmdText, fileUrl));
                try
                {
                    //打开连接
                    if (conn.State == ConnectionState.Broken || conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }

                    System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //获取Excel的第一个Sheet名称
                    string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim();
                    //查询sheet中的数据
                    string strSql = "select * from [" + sheetName + "]";
                    OleDbDataAdapter da = new OleDbDataAdapter(strSql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, table);
                    //ds.WriteXml(@"D:\xml.xml", XmlWriteMode.IgnoreSchema);  
                    dt = ds.Tables[0];
                    //ds.DataSetName = "sheshuo";
                    //dt.TableName = "Key";
                    //dt.WriteXml(@"D:\xml.xml", XmlWriteMode.IgnoreSchema);  
                    return dt;
                }
                catch (Exception exc)
                {
                    WriteLog.WriteError(exc.ToString());
                    throw exc;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            //private string strConnection = "Server=TTADMINISTRATOR;DataBase=qec_0;Uid=sa;pwd=sa;";
            //private string strConnection = "Server=.;DataBase=qec_0;Uid=tcg;pwd=tcg2011;";
            //private string strConnection = "Server=.;DataBase=qec_0;Uid=glsa;pwd=glsa201508;";
            private string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";

            /// <summary>
            /// 从System.Data.DataTable导入数据到数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int InsetData(System.Data.DataTable dt)
            {
                int i = 0;
                string lng = "";
                string lat = "";
                string offsetLNG = "";
                string offsetLAT = "";
                string strSql1 = "truncate table DBToExcel";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                //string strConnection = "Server=TTADMINISTRATOR;DataBase=qec_0;Uid=sa;pwd=sa;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    sqlConnection1.Close();
                }
                foreach (DataRow dr in dt.Rows)
                {
                    lng = dr["LNG"].ToString().Trim();
                    lat = dr["LAT"].ToString().Trim();
                    offsetLNG = dr["OFFSET_LNG"].ToString().Trim();
                    offsetLAT = dr["OFFSET_LAT"].ToString().Trim();
                    //sw = string.IsNullOrEmpty(sw) ? "null" : sw;
                    //kr = string.IsNullOrEmpty(kr) ? "null" : kr;
                    string strSql = string.Format("Insert into DBToExcel (LNG,LAT,OFFSET_LNG,OFFSET_LAT) Values ({0},'{1}','{2}','{3}')", lng, lat, offsetLNG, offsetLAT);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                    //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                    SqlConnection sqlConnection = new SqlConnection(strConnection);
                    try
                    {
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                        sqlConnection.Open();
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                    }
                    catch (Exception ex)
                    {
                        WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        sqlConnection.Close();
                    }
                
                }
                return i;
            }



            /// <summary>
            /// 从t_cardDataTable导入数据到t_card数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_cardInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                string card_id = "";
                string card_name = "";
                string card_quality = "";
                string card_HP = "";
                string card_atk = "";
                string card_level = "";
                string hp_growth = "";
                string atk_growth = "";
                string fight_growth = "";
                string card_desc = "";
                string card_race = "";
                string card_attr = "";
                string card_pic = "";
                string head_pic = "";
                string passivity_skill = "";
              
                string strSql1 = "truncate table t_card";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }

                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                    
                            card_id = dr["card_id"].ToString().Trim();
                            card_name = dr["card_name"].ToString().Trim();
                            card_quality = dr["card_quality"].ToString().Trim();
                            card_HP = dr["card_HP"].ToString().Trim();
                            card_atk = dr["card_atk"].ToString().Trim();
                            card_level = dr["card_level"].ToString().Trim();
                            hp_growth = dr["hp_growth"].ToString().Trim();
                            atk_growth = dr["atk_growth"].ToString().Trim();
                            fight_growth = dr["fight_growth"].ToString().Trim();
                            card_desc = dr["card_desc"].ToString().Trim();
                            card_race = dr["card_race"].ToString().Trim();
                            card_attr = dr["card_attr"].ToString().Trim();
                            card_pic = dr["card_pic"].ToString().Trim();
                            head_pic = dr["head_pic"].ToString().Trim();
                    
                   
                            string strSql = string.Format("Insert into t_card (card_id,card_name,card_quality,card_HP,card_atk,card_level,hp_growth,atk_growth,fight_growth,card_desc,card_race,card_attr,card_pic,head_pic,passivity_skill) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')", card_id,card_name,card_quality,card_HP,card_atk,card_level,hp_growth,atk_growth,fight_growth,card_desc,card_race,card_attr,card_pic,head_pic,passivity_skill);
                            //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                           //// string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                           
                    
                                // SqlConnection sqlConnection = new SqlConnection(strConnection);
                                
                                SqlCommand sqlCmd = new SqlCommand();
                                sqlCmd.CommandText = strSql;
                                sqlCmd.Connection = sqlConnection;
                                SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                                i++;
                                sqlDataReader.Close();
                    

                        }
                    }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + card_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }


            /// <summary>
            /// 从t_card_combDataTable导入数据到t_card_comb数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_card_combInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                string card_id = "";
                string material_id = "";
                string material_num = "";
                string need_gold = "";
                string success_rate = "";


                string strSql1 = "truncate table t_card_comb";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }
                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        card_id = dr["card_id"].ToString().Trim();
                        material_id = dr["material_id"].ToString().Trim();
                        material_num = dr["material_num"].ToString().Trim();
                        need_gold = dr["need_gold"].ToString().Trim();
                        success_rate = dr["success_rate"].ToString().Trim();



                        string strSql = string.Format("Insert into t_card_comb (card_id,material_id,material_num,need_gold,success_rate) Values ('{0}','{1}','{2}','{3}','{4}')", card_id, material_id, material_num, need_gold, success_rate);
                        //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                        //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                       
                      
                            // SqlConnection sqlConnection = new SqlConnection(strConnection);
                           
                            SqlCommand sqlCmd = new SqlCommand();
                            sqlCmd.CommandText = strSql;
                            sqlCmd.Connection = sqlConnection;
                            SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                            i++;
                            sqlDataReader.Close();
                       

                    }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + card_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }

            /// <summary>
            /// 从t_card_skillbDataTable导入数据到t_card_skill数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_card_skillInsetData(System.Data.DataTable dt)
            {
                int i = 0;
               
                string card_id = "";
                string skill_type = "";
                string skill_name = "";
                string description = "";



                string strSql1 = "truncate table t_card_skill";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    
                    card_id = dr["card_id"].ToString().Trim();
                    skill_type = dr["skill_type"].ToString().Trim();
                    skill_name = dr["skill_name"].ToString().Trim();
                    description = dr["description"].ToString().Trim();



                    string strSql = string.Format("Insert into t_card_skill (card_id,skill_type,skill_name,description) Values ('{0}','{1}','{2}','{3}')", card_id, skill_type, skill_name, description);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                    //string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                    
                      // SqlConnection sqlConnection = new SqlConnection(strConnection);
                       
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + card_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }

         
            /// <summary>
            /// 从t_chapter_eventbDataTable导入数据到t_chapter_event数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_chapter_eventInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                
                string chapter_id = "";
                string event_id = "";
                string event_rate = "";
                



                string strSql1 = "truncate table t_chapter_event";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               //// string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    
                    chapter_id = dr["chapter_id"].ToString().Trim();
                    event_id = dr["event_id"].ToString().Trim();
                    event_rate = dr["event_rate"].ToString().Trim();




                    string strSql = string.Format("Insert into t_chapter_event (chapter_id,event_id,event_rate) Values ('{0}','{1}','{2}')", chapter_id, event_id, event_rate);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                   //// string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                       
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + event_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }


            /// <summary>
            /// 从t_element_giftDataTable导入数据到t_element_gift数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_element_giftInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                string g_id = "";
                string g_type = "";
                string g_value = "";
                string g_sign = "";
                string g_num = "";
                string g_ratio = "";
                


                string strSql1 = "truncate table t_element_gift";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    g_id = dr["g_id"].ToString().Trim();
                    g_type = dr["g_type"].ToString().Trim();
                    g_value = dr["g_value"].ToString().Trim();
                    g_sign = dr["g_sign"].ToString().Trim();
                    g_num = dr["g_num"].ToString().Trim();
                    g_ratio = dr["g_ratio"].ToString().Trim();



                    string strSql = string.Format("Insert into t_element_gift (g_id,g_type,g_value,g_sign,g_num,g_ratio) Values ('{0}','{1}','{2}','{3}','{4}','{5}')", g_id, g_type, g_value, g_sign, g_num, g_ratio);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                    
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                       
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                  

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + g_value + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }

            /// <summary>
            /// 从t_element_shopDataTable导入数据到t_element_shop数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_element_shopInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                string shop_id = ""; 
                string commodity = "";
                string exp_type = ""; 
                string exp_num = ""; 
                string item_id = ""; 
                string item_num = ""; 
                string free_time = ""; 
                string discount_pic = ""; 
                string g_id = "";




                string strSql1 = "truncate table t_element_shop";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }




                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    shop_id = dr["shop_id"].ToString().Trim();
                    commodity = dr["commodity"].ToString().Trim();
                    exp_type = dr["exp_type"].ToString().Trim();
                    exp_num = dr["exp_num"].ToString().Trim();
                    item_id = dr["item_id"].ToString().Trim();
                    item_num = dr["item_num"].ToString().Trim();
                    free_time = dr["free_time"].ToString().Trim();
                    discount_pic = dr["discount_pic"].ToString().Trim();
                    g_id = dr["g_id"].ToString().Trim();



                    string strSql = string.Format("Insert into t_element_shop (shop_id	,commodity,	exp_type	,exp_num	,item_id	,item_num,	free_time,	discount_pic,	g_id) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", shop_id, commodity, exp_type, exp_num, item_id, item_num, free_time, discount_pic, g_id);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                   
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                      
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + shop_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }


            /// <summary>
            /// 从t_eventDataTable导入数据到t_event数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_eventInsetData(System.Data.DataTable dt)
            {
                int i = 0;
                
                string event_id = ""; 
                string event_name = ""; 
                string event_aim = ""; 
                string event_desc = "";
                string event_type = ""; 
                string event_finish_type = ""; 
                string event_isOpen = ""; 
                string event_trigger = ""; 
                string event_finish_need = ""; 
                string event_rewards = ""; 
                string event_canGiveup = ""; 
                string event_level = ""; 
                string event_rate = ""; 
                string event_pic = "";



                string strSql1 = "truncate table t_event";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    event_id = dr["event_id"].ToString().Trim();
                    event_name = dr["event_name"].ToString().Trim();
                    event_aim = dr["event_aim"].ToString().Trim();
                    event_desc = dr["event_desc"].ToString().Trim();
                    event_type = dr["event_type"].ToString().Trim();
                    event_finish_type = dr["event_finish_type"].ToString().Trim();
                    event_isOpen = dr["event_isOpen"].ToString().Trim();
                    event_trigger = dr["event_trigger"].ToString().Trim();
                    event_finish_need = dr["event_finish_need"].ToString().Trim();
                    event_rewards = dr["event_rewards"].ToString().Trim();
                    event_canGiveup = dr["event_canGiveup"].ToString().Trim();
                    event_level = dr["event_level"].ToString().Trim();
                    event_rate = dr["event_rate"].ToString().Trim();
                    event_pic = dr["event_pic"].ToString().Trim();




                    string strSql = string.Format("Insert into t_event (event_id,event_name,event_aim,event_desc,event_type,event_finish_type,event_isOpen,event_trigger,event_finish_need,event_rewards,event_canGiveup,event_level,event_rate,event_pic) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", event_id, event_name, event_aim, event_desc, event_type, event_finish_type, event_isOpen, event_trigger, event_finish_need, event_rewards, event_canGiveup, event_level, event_rate, event_pic);
                    //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                    
                   
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                      
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + event_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }


            /// <summary>
            /// 从t_init_card_groupDataTable导入数据到t_init_card_group数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_init_card_groupInsetData(System.Data.DataTable dt)
            {
                int i = 0;

                string group_id = "";
                string card_id = "";




                string strSql1 = "truncate table t_init_card_group";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                   
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                 SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {

                    group_id = dr["group_id"].ToString().Trim();
                    card_id = dr["card_id"].ToString().Trim();




                    string strSql = string.Format("Insert into t_init_card_group (group_id,card_id) Values ('{0}','{1}')", group_id, card_id);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                   
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                       
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                  

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + card_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }



            /// <summary>
            /// 从t_itemDataTable导入数据到t_item数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_itemInsetData(System.Data.DataTable dt)
            {
                int i = 0;

                
                string item_id = "";
                string item_name = "";
                string item_type = "";
                string item_desc = "";
                string type_desc = "";
                string item_value = "";
                string item_quality	= "";
                string item_recycle	= "";
                string map_id = "";




                string strSql1 = "truncate table t_item";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }




                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    item_id = dr["item_id"].ToString().Trim();
                    item_name = dr["item_name"].ToString().Trim();
                    item_type = dr["item_type"].ToString().Trim();
                    item_desc = dr["item_desc"].ToString().Trim();
                    type_desc = dr["type_desc"].ToString().Trim();
                    item_value = dr["item_value"].ToString().Trim();
                    item_quality = dr["item_quality"].ToString().Trim();
                    item_recycle = dr["item_recycle"].ToString().Trim();
                    map_id = dr["map_id"].ToString().Trim();




                    string strSql = string.Format("Insert into t_item (item_id,item_name,item_type,item_desc,type_desc,item_value,item_quality,item_recycle,map_id) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", item_id,item_name,item_type,item_desc,type_desc,item_value,item_quality,item_recycle,map_id);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        // SqlConnection sqlConnection = new SqlConnection(strConnection);
                        
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   
                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + item_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }

            /// <summary>
            /// 从t_item_combDataTable导入数据到t_item_comb数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_item_combInsetData(System.Data.DataTable dt)
            {
                int i = 0;


                string item_id = "";
                string material_id = ""; 
                string material_num = ""; 
                string m_ratio = "";





                string strSql1 = "truncate table t_item_comb";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    item_id = dr["item_id"].ToString().Trim();
                    material_id = dr["material_id"].ToString().Trim();
                    material_num = dr["material_num"].ToString().Trim();
                    m_ratio = dr["m_ratio"].ToString().Trim();




                    string strSql = string.Format("Insert into t_item_comb (item_id,material_id,material_num,m_ratio) Values ('{0}','{1}','{2}','{3}')", item_id, material_id, material_num, m_ratio);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + item_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }



            /// <summary>
            /// 从t_item_enlargeDataTable导入数据到t_item_enlarge数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_item_enlargeInsetData(System.Data.DataTable dt)
            {
                int i = 0;


                string add_num = "";
                string need_gold = "";
                string need_item = "";




                string strSql1 = "truncate table t_item_enlarge";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    add_num = dr["add_num"].ToString().Trim();
                    need_gold = dr["need_gold"].ToString().Trim();
                    need_item = dr["need_item"].ToString().Trim();





                    string strSql = string.Format("Insert into t_item_enlarge (add_num,need_gold,need_item) Values ('{0}','{1}','{2}')", add_num, need_gold, need_item);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                    
                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + add_num + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }




            /// <summary>
            /// 从t_mapDataTable导入数据到t_map数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_mapInsetData(System.Data.DataTable dt)
            {
                int i = 0;


                string map_id = "";
                string map_name = "";
                string min_level = "";
                string max_level = "";
                string map_desc = "";
                string map_group_id = "";
                string next_map = "";
                string event_odds = "";
                string map_level = "";




                string strSql1 = "truncate table t_map";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    map_id = dr["map_id"].ToString().Trim();
                    map_name = dr["map_name"].ToString().Trim();
                    min_level = dr["min_level"].ToString().Trim();
                    max_level = dr["max_level"].ToString().Trim();
                    map_desc = dr["map_desc"].ToString().Trim();
                    map_group_id = dr["map_group_id"].ToString().Trim();
                    next_map = dr["next_map"].ToString().Trim();
                    event_odds = dr["event_odds"].ToString().Trim();
                    map_level = dr["map_level"].ToString().Trim();





                    string strSql = string.Format("Insert into t_map (map_id,map_name,min_level,max_level,map_desc,map_group_id,next_map,event_odds,map_level) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", map_id, map_name, min_level, max_level, map_desc, map_group_id, next_map, event_odds, map_level);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                    
                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + map_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }





            /// <summary>
            /// 从t_map_groupDataTable导入数据到t_map_group数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_map_groupInsetData(System.Data.DataTable dt)
            {
                int i = 0;


                string map_group_id = "";
                string map_group_name = "";
                string need_main = "";
                string t_map_chapter = "";
                string PosXX = "";
                string PosYY = "";



                string strSql1 = "truncate table t_map_group";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    map_group_id = dr["map_group_id"].ToString().Trim();
                    map_group_name = dr["map_group_name"].ToString().Trim();
                    need_main = dr["need_main"].ToString().Trim();
                    t_map_chapter = dr["t_map_chapter"].ToString().Trim();
                    PosXX = dr["PosXX"].ToString().Trim();
                    PosYY = dr["PosYY"].ToString().Trim();






                    string strSql = string.Format("Insert into t_map_group (map_group_id,map_group_name,need_main,screen_id,PosXX,PosYY) Values ('{0}','{1}','{2}','{3}','{4}','{5}')", map_group_id, map_group_name, need_main, t_map_chapter, PosXX, PosYY);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                    

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + map_group_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }




            /// <summary>
            /// 从t_map_monsterDataTable导入数据到t_map_monster数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_map_monsterInsetData(System.Data.DataTable dt)
            {
                int i = 0;


                string map_id = "";
                string monster_id = "";
                string mm_rate = "";
                string reward_name = "";
               
                


                string strSql1 = "truncate table t_map_monster";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                 SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    map_id = dr["map_id"].ToString().Trim();
                    monster_id = dr["monster_id"].ToString().Trim();
                    mm_rate = dr["mm_rate"].ToString().Trim();
                    reward_name = dr["reward_name"].ToString().Trim();
                   





                    string strSql = string.Format("Insert into t_map_monster (map_id,monster_id,mm_rate,reward_name) Values ('{0}','{1}','{2}','{3}')", map_id, monster_id, mm_rate, reward_name);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + map_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }




            /// <summary>
            /// 从t_monsterDataTable导入数据到t_monster数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_monsterInsetData(System.Data.DataTable dt)
            {
                int i = 0;

                string monster_id = "";
                string monster_name = "";
                string monster_desc = "";
                string monster_type = "";
                string monster_levle = "";
                string crit_rate = "";
                string hit_rate = "";
                string evade_rate = "";
                string head_pic = "";
                
                
                
                
                
                
                
                



                string strSql1 = "truncate table t_monster";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                   
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                    foreach (DataRow dr in dt.Rows)
                    {


                        monster_id = dr["monster_id"].ToString().Trim();
                        monster_name = dr["monster_name"].ToString().Trim();
                        monster_desc = dr["monster_desc"].ToString().Trim();
                        monster_type = dr["monster_type"].ToString().Trim();
                        monster_levle = dr["monster_levle"].ToString().Trim();
                        crit_rate = dr["crit_rate"].ToString().Trim();
                        hit_rate = dr["hit_rate"].ToString().Trim();
                        evade_rate = dr["evade_rate"].ToString().Trim();
                        head_pic = dr["head_pic"].ToString().Trim();






                        string strSql = string.Format("Insert into t_monster (monster_id,monster_name,monster_desc,monster_type,monster_level,crit_rate,hit_rate,evade_rate,head_pic) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", monster_id, monster_name, monster_desc, monster_type, monster_levle, crit_rate, hit_rate, evade_rate, head_pic);
                        // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";

                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();

                    }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + monster_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }





            /// <summary>
            /// 从t_monster_cardDataTable导入数据到t_monster_card数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_monster_cardInsetData(System.Data.DataTable dt)
            {
                int i = 0;

           
                string monster_id = "";
                string card_id = "";
                string card_name = "";
                string card_nick = ""; 
                string card_level = ""; 
                string card_status = ""; 
                string card_exp = ""; 
                string intensify_level = "";
                string card_quality = "";
                string card_HP = ""; 
                string card_atk = ""; 
                string hp_growth = ""; 
                string atk_growth = ""; 
                string fight_growth = ""; 
                string card_desc = ""; 
                string card_race = ""; 
                string card_attr = ""; 
                string card_addTime = ""; 
                string phase_exp = ""; 
                string card_pic = ""; 
                string head_pic = "";









                string strSql1 = "truncate table t_monster_card";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                  
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                 SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    monster_id = dr["monster_id"].ToString().Trim();
                    card_id = dr["card_id"].ToString().Trim();
                    card_name = dr["card_name"].ToString().Trim();
                    card_nick = dr["card_nick"].ToString().Trim();
                    card_level = dr["card_level"].ToString().Trim();
                    card_status = dr["card_status"].ToString().Trim();
                    card_exp = dr["card_exp"].ToString().Trim();
                    intensify_level = dr["intensify_level"].ToString().Trim();
                    card_quality = dr["card_quality"].ToString().Trim();
                    card_HP = dr["card_HP"].ToString().Trim();
                    card_atk = dr["card_atk"].ToString().Trim();
                    hp_growth = dr["hp_growth"].ToString().Trim();
                    atk_growth = dr["atk_growth"].ToString().Trim();
                    fight_growth = dr["fight_growth"].ToString().Trim();
                    card_desc = dr["card_desc"].ToString().Trim();
                    card_race = dr["card_race"].ToString().Trim();
                    card_attr = dr["card_attr"].ToString().Trim();
                    card_addTime = dr["card_addTime"].ToString().Trim();
                    phase_exp = dr["phase_exp"].ToString().Trim();
                    card_pic = dr["card_pic"].ToString().Trim();
                    head_pic = dr["head_pic"].ToString().Trim();






                    string strSql = string.Format("Insert into t_monster_card (monster_id,card_id,card_name,card_nick,card_level,card_status,card_exp,intensify_level,card_quality,card_HP,card_atk,hp_growth,atk_growth,fight_growth,card_desc,card_race,card_attr,card_addTime,phase_exp,card_pic,head_pic) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}')", monster_id, card_id, card_name, card_nick, card_level, card_status, card_exp, intensify_level, card_quality, card_HP, card_atk, hp_growth, atk_growth, fight_growth, card_desc, card_race, card_attr, card_addTime, phase_exp, card_pic, head_pic);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   
                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + monster_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }








            /// <summary>
            /// 从t_monster_rewardDataTable导入数据到t_monster_reward数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_monster_rewardInsetData(System.Data.DataTable dt)
            {
                int i = 0;

                string rg_id = ""; 
                string r_type = ""; 
                string r_value = ""; 
                string r_num = ""; 
                string r_ratio = "";
                








                string strSql1 = "truncate table t_monster_reward";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                   
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    rg_id = dr["rg_id"].ToString().Trim();
                    r_type = dr["r_type"].ToString().Trim();
                    r_value = dr["r_value"].ToString().Trim();
                    r_num = dr["r_num"].ToString().Trim();
                    r_ratio = dr["r_ratio"].ToString().Trim();
                   






                    string strSql = string.Format("Insert into t_monster_reward (rg_id,r_type,r_value,r_num,r_ratio) Values ('{0}','{1}','{2}','{3}','{4}')", rg_id, r_type, r_value, r_num, r_ratio);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + rg_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }




            /// <summary>
            /// 从t_pvp_shopDataTable导入数据到t_pvp_shop数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_pvp_shopInsetData(System.Data.DataTable dt)
            {
                int i = 0;

               
                string s_id = "";
                string s_name = ""; 
                string s_type = ""; 
                string v_type = ""; 
                string s_value = "";
                string s_num = "";
                string s_price = ""; 
                string s_rate = "";

                






                string strSql1 = "truncate table t_pvp_shop";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    s_id = dr["s_id"].ToString().Trim();
                    s_name = dr["s_name"].ToString().Trim();
                    s_type = dr["s_type"].ToString().Trim();
                    v_type = dr["v_type"].ToString().Trim();
                    s_value = dr["s_value"].ToString().Trim();
                    s_num = dr["s_num"].ToString().Trim();
                    s_price = dr["s_price"].ToString().Trim();
                    s_rate = dr["s_rate"].ToString().Trim();
                   






                    string strSql = string.Format("Insert into t_pvp_shop (s_id,s_name,s_type,v_type,s_value,s_num,s_price,s_rate) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", s_id, s_name, s_type, v_type, s_value, s_num, s_price, s_rate);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + s_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }





            /// <summary>
            /// 从t_reward_groupDataTable导入数据到t_reward_group数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_reward_groupInsetData(System.Data.DataTable dt)
            {
                int i = 0;


               
                string id = "";
                string group_name = "";
                string reward_rate = "";
                string is_win = "";







                string strSql1 = "truncate table t_reward_group";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
               // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                   
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                finally
                {
                    sqlConnection1.Close();
                }



                 SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    id = dr["id"].ToString().Trim();
                    group_name = dr["group_name"].ToString().Trim();
                    reward_rate = dr["reward_rate"].ToString().Trim();
                    is_win = dr["is_win"].ToString().Trim();
                    







                    string strSql = string.Format("Insert into t_reward_group (id,group_name,reward_rate,is_win) Values ('{0}','{1}','{2}','{3}')", id, group_name, reward_rate, is_win);
                   // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }





            /// <summary>
            /// 从t_vipDataTable导入数据到t_vip数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_vipInsetData(System.Data.DataTable dt)
            {
                int i = 0;



                string vip_id = "";
                string topUp_num = "";
                string mopUp = "";
                string buy_main = "";
                string buy_gold = "";
                string reset_boss = "";
                string buy_athletics = "";
                string speed_onHook = "";
                string unlock_boss = "";
                string offLine = "";
                string unlock_ath_cd = "";
                string unlock_skip_ath = "";
                string unlock_random_gold = "";
                string dispatch_time_reduction = "";
                string accelerate_dispatch = "";
                string boss_gold = "";
                string random_event_sum = "";
                string dispatch_gold_add = "";
                string event_refresh = "";
                string dispatch_meanwhile_num = "";
                string unlock_finish_dispatch = "";
                string unlock_finish_all_random = "";







                string strSql1 = "truncate table t_vip";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                   
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection1.Close();
                }


                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    vip_id = dr["vip_id"].ToString().Trim();
                    topUp_num = dr["topUp_num"].ToString().Trim();
                    mopUp = dr["mopUp"].ToString().Trim();
                    buy_main = dr["buy_main"].ToString().Trim();
                    buy_gold = dr["buy_gold"].ToString().Trim();
                    reset_boss = dr["reset_boss"].ToString().Trim();
                    buy_athletics = dr["buy_athletics"].ToString().Trim();
                    speed_onHook = dr["speed_onHook"].ToString().Trim();
                    unlock_boss = dr["unlock_boss"].ToString().Trim();
                    offLine = dr["offLine"].ToString().Trim();
                    unlock_ath_cd = dr["unlock_ath_cd"].ToString().Trim();
                    unlock_skip_ath = dr["unlock_skip_ath"].ToString().Trim();
                    unlock_random_gold = dr["unlock_random_gold"].ToString().Trim();
                    dispatch_time_reduction = dr["dispatch_time_reduction"].ToString().Trim();
                    accelerate_dispatch = dr["accelerate_dispatch"].ToString().Trim();
                    boss_gold = dr["boss_gold"].ToString().Trim();
                    random_event_sum = dr["random_event_sum"].ToString().Trim();
                    dispatch_gold_add = dr["dispatch_gold_add"].ToString().Trim();
                    event_refresh = dr["event_refresh"].ToString().Trim();
                    dispatch_meanwhile_num = dr["dispatch_meanwhile_num"].ToString().Trim();
                    unlock_finish_dispatch = dr["unlock_finish_dispatch"].ToString().Trim();
                    unlock_finish_all_random = dr["unlock_finish_all_random"].ToString().Trim();









                    string strSql = string.Format("Insert into t_vip (vip_id,topUp_num,mopUp,buy_main,buy_gold,reset_boss,buy_athletics,speed_onHook,unlock_boss,offLine,unlock_ath_cd,unlock_skip_ath,unlock_random_gold,dispatch_time_reduction,accelerate_dispatch,boss_gold,random_event_sum,dispatch_gold_add,event_refresh,dispatch_meanwhile_num,unlock_finish_dispatch,unlock_finish_all_random) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}')", vip_id, topUp_num, mopUp, buy_main, buy_gold, reset_boss, buy_athletics, speed_onHook, unlock_boss, offLine, unlock_ath_cd, unlock_skip_ath, unlock_random_gold, dispatch_time_reduction, accelerate_dispatch, boss_gold, random_event_sum, dispatch_gold_add, event_refresh, dispatch_meanwhile_num, unlock_finish_dispatch, unlock_finish_all_random);
                    // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                    

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + vip_id + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }




            /// <summary>
            /// 从t_userLevel_confDataTable导入数据到t_userLevel_conf数据库
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public int t_userLevel_confInsetData(System.Data.DataTable dt)
            {
                int i = 0;



                string levelNum = ""; 
                string exp = ""; 
                string cardNum = ""; 
                string cardGroupNum = "";
                string fight = ""; 
                string crit = "";
                string hit = "";
                string evade = "";
                string cardBag = "";
                string itemBag = "";
                string main = "";
                





                string strSql1 = "truncate table t_userLevel_conf";
                //string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
                // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                SqlConnection sqlConnection1 = new SqlConnection(strConnection);
                try
                {
                    // SqlConnection sqlConnection = new SqlConnection(strConnection);
                    sqlConnection1.Open();
                    SqlCommand sqlCmd = new SqlCommand();
                    sqlCmd.CommandText = strSql1;
                    sqlCmd.Connection = sqlConnection1;
                    SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                    
                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError(ex.ToString());DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection1.Close();
                }



                SqlConnection sqlConnection = new SqlConnection(strConnection);
                sqlConnection.Open();
                try
                {
                foreach (DataRow dr in dt.Rows)
                {


                    levelNum = dr["levelNum"].ToString().Trim();
                    exp = dr["exp"].ToString().Trim();
                    cardNum = dr["cardNum"].ToString().Trim();
                    cardGroupNum = dr["cardGroupNum"].ToString().Trim();
                    fight = dr["fight"].ToString().Trim();
                    crit = dr["crit"].ToString().Trim();
                    hit = dr["hit"].ToString().Trim();
                    evade = dr["evade"].ToString().Trim();
                    cardBag = dr["cardBag"].ToString().Trim();
                    itemBag = dr["itemBag"].ToString().Trim();
                    main = dr["main"].ToString().Trim();
                   









                    string strSql = string.Format("Insert into t_userLevel_conf (levelNum,exp,cardNum,cardGroupNum,fight,crit,hit,evade,cardBag,itemBag,main) Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", levelNum, exp, cardNum, cardGroupNum, fight, crit, hit, evade, cardBag, itemBag, main);
                    // string strConnection = "Server=.;DataBase=HW;Uid=admin;pwd=admin;";
                   
                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.CommandText = strSql;
                        sqlCmd.Connection = sqlConnection;
                        SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
                        i++;
                        sqlDataReader.Close();
                   

                }
                }
                catch (Exception ex)
                {
                    WriteLog.WriteError("出错ID：" + levelNum + "|" + ex.ToString()); DialogResult error = MessageBox.Show(ex.Message.ToString(), "出错！", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                finally
                {
                    sqlConnection.Close();
                }
                return i;
            }
    }
}
