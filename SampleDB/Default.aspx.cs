using System;
using System.Web;
using System.Web.UI;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

/*
 * 参考URL　https://www.youtube.com/watch?v=fJu8Zo-q25Q
 * 動画名 asp.net c# bind records in gridview between two dates sql server
 */
namespace SampleDB
{

    public partial class Default : System.Web.UI.Page
    {
        public void button1Clicked(object sender, EventArgs args)
        {
            button1.Text = "You clicked me";
            string Text = "任意の値";
            //DB接続情報を取得
            string connectionString = ConfigurationManager.ConnectionStrings[1].ToString();

            //ネットワーク的な接続を開始
            using(SqlConnection con = new SqlConnection(connectionString))
            {
                //テーブルからデータを受けとる為に用意する変数
                SqlDataReader sqlDataReader = null;

                //DB接続開始
                con.Open();
                try
                {
                    //SQL文の設定
                    string strSql = "select * from SampleDB where ProductName = @ProductName";

                    //接続したデータベースに対するSQL文の予約
                    SqlCommand command = new SqlCommand(strSql, con);

                    //パラメータの設定
                    command.Parameters.Add(new SqlParameter("@ProductName", Text));

                    /*
                     * 以下の①から③研修でやっていない内容
                     * DataAdapterとは、その名前が表すようにデータベース（SQLサーバ内のDB）とデータセット（GridViewに表示させたいテーブル）
                     * の間をつなぐ「アダプタ」の役目を果たす
                     * 参考URL　http://www.atmarkit.co.jp/ait/articles/0309/06/news002.html
                     * ①アダプターにSQLサーバから取って来たデータを入れる
                     * ②ローカルのデータテーブルを作成
                     * ③Fillメソッドを使い、データテーブルにアダプターの内容を入れる
                     */
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);//①

                    DataTable dataTable = new DataTable();//②

                    sqlDataAdapter.Fill(dataTable);//③

                    //読み取り実行でSQLを実行
                    sqlDataReader = command.ExecuteReader();

                    if(sqlDataReader.Read())
                    {
                        //GridViewにデータを反映させる
                        Grid1.DataSource = dataTable;
                        Grid1.DataBind();
                    }
                }
                catch(Exception ex)
                {
                    Response.Write(ex.Message);
                }
                finally
                {
                    sqlDataReader.Close();
                }
            }
        }
    }
}
