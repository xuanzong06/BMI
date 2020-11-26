https://www.c-sharpcorner.com/UploadFile/6b8651/read-excel-file-in-windows-application-using-C-Sharp/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable ReadExcel(string FileName, string FileExt)
        {
            string Source = string.Empty;
            DataTable dtexcewl = new DataTable();
            if(FileExt.CompareTo(".xls") == 0){
                Source = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007 
            }
            else {
                Source = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  
            }

            using(OleDbConnection conn = new OleDbConnection(Source)){
                try
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [10908$]", conn);
                    adapter.Fill(dtexcewl);
                }
                catch (Exception ex)
                {
                    string ex_txt = ex.Message.ToString();
                }
            }
            return dtexcewl;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if(file.ShowDialog() == System.Windows.Forms.DialogResult.OK){
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt);
                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex != 0)
                {
                    MessageBox.Show("請點擊【第一個】欄位");
                }
                else
                {
                    label1.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    label2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+1].Value.ToString();
                    label3.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+2].Value.ToString();
                    label4.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+3].Value.ToString();
                    label5.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+4].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤訊息：" +ex.Message.ToString());
            }
        }

       
    }
}


# BMI

Module 22. 認識資料庫異動
22.1	使用SQL敘述控制單一資料庫異動
22.2	使用ADO.NET控制單一資料庫異動
22.3	分散式交易
參考資料:https://docs.microsoft.com/zh-tw/sql/relational-databases/track-changes/about-change-data-capture-sql-server?view=sql-server-ver15

目錄
1.本章來認識資料庫異動(資料庫被更動時會有什麼事發生)
2.當我們對資料庫進行資料修改(新增、刪除、更新)的時候，資料庫會紀錄我們所做的活動
3.所以如果有要查詢異動的需求時，可以找到異動的詳細資料
4.本章會帶大家認識異動資料擷取的資料流程
5.理解異動資料擷取和傑取執行個體

22.1	使用SQL敘述控制單一資料庫異動
1.異動資料擷取會記錄套用至 SQL Server 資料表的插入、更新和刪除活動
2.這樣會以方便取用的「關聯式格式」提供變更的詳細資料
3.系統會針對修改的資料列擷取資料行資訊以及將變更套用至目標環境所需的中繼資料
4.並且將它們儲存在鏡像追蹤來源資料表的資料行結構的變更資料表中
5.除此之外，系統會提供資料表值函式，讓取用者以有系統的方式存取異動資料

(二)
1.此技術之目標資料取用者的理想範例為擷取、轉換和下載 (ETL) 應用程式
2.ETL 應用程式會將變更資料從 SQL Server 來源資料表累加地載入資料倉儲
3.雖然在資料倉儲內的來源資料表表示法必須反映來源資料表中的變更
4.但是重新整理來源複本的端對端技術並不適用
5.SQL Server 異動資料擷取提供結構化變更資料的可靠資料流，讓取用者可以將其套用到不同的資料目標表示法

(三)
1.異動資料擷取的變更資料來源是 SQL Server 交易記錄
2.當插入、更新和刪除作業套用至追蹤來源資料表時，描述這些變更的項目就會加入記錄
3.這個記錄會被當做擷取程序的輸入→這樣就會讀取記錄並將變更之相關資訊新增至追蹤資料表的相關聯變更資料表
4.系統會提供一些函數，用來列舉指定的範圍內出現在變更資料表中的變更，並以篩選結果集的形式傳回此資訊
5.應用程式處理序通常會使用篩選結果集，在某些外部環境中更新來源的表示法。

(四)
1.所以我們了解到，當我們對資料庫中的資料進行更動，所做的行為都會被資料庫記錄
2.因為所做的行為都被記錄，所以每個動作都是可以被追溯的→用來找出造成異常的動作
3.SQL Server 交易紀錄指的是資料庫應用程式與資料庫伺服器之間的作業，每當成功執行改變資料庫的內容時，就是一個成功的交易(應用程式給資料庫，資料庫給應用程式)
4.每一筆完成的交易都會被記錄下來，我們可以使用這些紀錄去追蹤那些變更的資料表
5.資料庫異動的技術需要不斷的練習才能理解、熟用

參考資料:https://docs.microsoft.com/zh-tw/dotnet/framework/data/adonet/transactions-and-concurrency
22.2使用ADO.NET控制單一資料庫異動
1.異動是由單一命令或當做封裝 (Package) 執行的命令群組所組成
2.「交易」可讓我們開發時將多項作業結合成單一工作單位
3.如果異動的某一處失敗了，則所有更新都會復原到異動之前的狀態(保護的機制)
4.異動必須符合 ACID 屬性，分別是： 單元性 (Atomicity)、一致性 (Consistency)、隔離性 (Isolation) 和持續性 (Durability)，才能保證資料一致性
5.大多數關聯式資料庫系統 (例如 Microsoft SQL Server) 都可以支援異動，其方法是在每次用戶端應用程式執行更新、插入或刪除作業時，提供鎖定、記錄和異動管理功能

（二）ADO.NET本機異動
1.當我們想要將多個工作在一起執行，讓程式以單一工作單位的形式執行時，會使用 ADO.NET 中的「交易」
2.案例發想→請想像應用程式正在執行兩項工作↓
3.第一:更新包含訂單資訊的資料表
4.第二:第一個工作完成後，更新包含存貨資訊的資料表，並將訂購的項目加在購買者身上
5.當第一或第二其中一項工作失敗的時候，就會自動地復原這兩個更新

（三）ADO.NET本機異動→決定異動的類型
1.當交易為單一階段交易時，ADO.NET就會將單一階段交易的交易視為本機交易，並直接由資料庫處理
2.每個 .NET Framework 的DataProvider 都有自己Transaction的物件(Sql有自己的，OleDB也有自己的)可以用來執行本機交易
3.我們操作的資料庫伺服器是SQL Server，因此我們要在SQL Server資料庫中執行異動時，請在開發的專案中，引用命名空間 System.Data.SqlClient
4.引用完命名空間，我們還有DbTransaction類別可以使用，可以用來撰寫需要交易
5.注意:當交易在伺服器上執行時，最有效率；以SQL Server來說，使用Transact-SQL或將作業寫入預存程序執行可以達到比較好的作業效率

（四）ADO.NET 搭配 Microsoft SQL Server 的交易式邏輯 程式範例
1.建立與SQL Server資料庫的連線，以及SqlTransaction物件 + sqlcommand物件
2.建立兩個command.CommandText執行insert的作業(故意將第二個insert作業去加入不存在的欄位，我們就可以觸發錯誤XD)
3.建立try...catch區塊，攔截例外，當例外發生的時候，我們可以使用 sqlTransaction 執行 RollBack
4.即使是寫在catch區塊的作業，也可以在加入一個try ... catch來攔截例外，就擔心執行RollBack時，遇到例外的狀況
5.(這裡檢查資料庫的資料~~~~~~~~~~~~)

22.3	分散式交易
1.異動是「一組」相關工作，意思是一個異動包含了作業的「成功→異動資料」或是「失敗→復原資料」
2.分散式交易是影響數個資源的交易→白話的意思是:系統存取的資料庫伺服器超過1台
3.因為分散式交易是影響數個資源的交易，所以有些規則必須遵守
4.系統對於要認可的分散式異動，所有參與者都必須保證資料的任何變更都是永久的
5.如果有任何參與者無法保證資料的變更是永久的，整個異動的作業都會失敗，並將異動的資料復原

(一)使用System.Transactions
1.在 .NET Framework 中，會透過 System.Transactions 命名空間中的 API 來管理分散式異動
2.當要連接到多個資料庫伺服器時，System.Transactions API 會將分散式異動處理委派給異動監視器
3.異動監視器→Microsoft 分散式異動協調器 (MS DTC)
4.ADO.NET 2.0 支援使用 EnlistTransaction 方法在分散式交易中登記，該方法會在 Transaction 執行個體中登記連接
5.交易的基礎觀念可以閱讀 MSDN:https://docs.microsoft.com/zh-tw/dotnet/framework/data/transactions/transaction-fundamentals

(二)自動登記分散式異動
1.自動登記是整合 ADO.NET 連接與 System.Transactions 的預設 (及慣用) 方式
2.如果連接物件確定交易處於作用中，就會自動在現有分散式交易中登記
3.作用中的定義:在 System.Transaction 詞彙中，表示 Transaction.Current 不為 Null
4.請注意:當開啟連接時，即會發生自動異動登記
5.因為自動地登記異動，所以之後即使在異動範圍內執行命令，也不會觸發自動異動登記的事件

(三)在分散式異動中手動登記
1.如果沒有啟用自動登記，或者是開發者有需要登記在開啟連線之後啟動的交易時，我們可以選擇手動登記
2.在這個情況:可以針對所使用的提供者，使用 EnlistTransaction 物件的 DbConnection 方法
3.在現有的分散式異動中登記可確保，如果認可(更新資料)或復原異動，則也會認可或復原資料來源的程式碼所做的修改
4.當共用商務物件時，在分散式異動中登記會特別適用
5.如果使用共用的商務物件來執行多個異動，則該物件的開啟連接將不會自動在新起始的異動中登記

參考資料:https://ithelp.ithome.com.tw/articles/10162454
(四) 在我的SQL Server中建立第二個資料庫
※我在思考這一part是要繼續理論下一章才實作嗎?
1.
2.
3.
4.
5.

※如果選擇理論:
(四)EnlistTransaction
1.EnlistTransaction採用類型Transaction的單一引數，這是現有交易的參考
2.呼叫連接的 EnlistTransaction 方法之後，對使用連接之資料來源所做的所有修改都會包含在異動中
3.傳遞 Null 值可將連接從目前的分散式異動登記中取消登記
4.在在呼叫 EnlistTransaction 之前，必須先開啟連接→Connection.Open();
5.注意:在交易上明確地登記連接之後，直到第一個交易完成之前，無法取消登記或在其他交易中登記它


分散式:當交易由交易監視器協調並使用不安全的機制（例如兩階段認可）進行交易解析時，就會將其視為分散式交易
參考一下:
120.105.184.250 › lwcheng › kid51 › kidpps › kid51_chap14
keyword : 分散式異動 範例

*******************************************************************************************************************************
Module 23. 資料庫異動控制實作
23.1	TransactionScope
23.2	CommittableTransaction 
23.3	DTC 分散式交易
參考資料:

目錄
1.
2.
3.
4.
5.

23.1 TransactionScope 類別
1.TransactionScope 類別→使程式碼區塊成為異動式。 這個類別無法被繼承。
2.System.Transactions 基礎結構同時提供以 Transaction 類別為基礎的明確程式設計模型
3.以及使用 TransactionScope 類別的隱含程式設計模型，其中的交易會由基礎結構自動管理
4.當 new 語句具現化 TransactionScope 時，交易管理員會決定要參與哪個交易
5.一旦決定後，範圍永遠會參與該異動。

(二)
1.此決策是根據兩個因素而定：環境異動是否存在，以及建構函式中的 TransactionScopeOption 參數值
2.環境交易是您的程式碼執行所在的交易
3.您可以呼叫 Transaction.Current 類別的靜態 Transaction 屬性，取得環境交易的參考
4.如果交易範圍內發生例外狀況→會回復它所參與的交易
5.如果交易範圍內未發生例外狀況→則允許範圍所參與的交易繼續進行

(三)
1.當程式完成在交易中執行的所有工作時，要設計呼叫 Complete 方法一次(只能一次)
2.告知交易管理員可接受認可交易。 如果無法呼叫這個Complete方法，就會中止交易
3.呼叫 Dispose 方法會標記交易範圍的結尾
4.在呼叫Dispose方法後發生的例外狀況 通常不太可能會影響異動。
5.

參考https://docs.microsoft.com/zh-tw/dotnet/api/system.transactions.transactionscope?view=netframework-4.8
(四)程式範例
1.
2.
3.
4.
5.

23.2 CommittableTransaction
1.CommittableTransaction 類別→描述可認可的交易
2.CommittableTransaction 類別為應用程式提供使用交易的明確方式，而非隱含地使用 TransactionScope 類別
3.通常會用 TransactionScope 類別來建立隱含交易，以便自動管理環境交易內容
4.建立 CommittableTransaction 不會自動設定環境交易
5.要取得或是設定環境需要呼叫全域 Transaction 物件的靜態 Transaction.Current 屬性來取得

參考連結:https://docs.microsoft.com/zh-tw/dotnet/api/system.transactions.committabletransaction?view=netframework-4.8
(二)建構函式
1.
2.
3.
4.
5.

參考連結:https://docs.microsoft.com/zh-tw/dotnet/framework/data/transactions/implementing-an-explicit-transaction-using-committabletransaction
(三)使用 CommittableTransaction 實作明確交易
1.CommittableTransaction 類別係衍生自 Transaction 類別，因此可提供Transaction 類別的所有功能
2.Rollback 類別上的 Transaction 方法特別有用，因為它同時能用來復原 CommittableTransaction 物件
3.Transaction 類別與 CommittableTransaction 類別類似，但是不會提供 Commit 方法
4.它能讓您在控制交易認可時間的同時，將交易物件 (或其複製品) 傳遞給其他方法 (可能透過其他執行緒)
5.呼叫的程式碼可以登記並投票給交易，但是只有 CommittableTransaction 物件的建立者有能力認可交易。

參考連結:https://docs.microsoft.com/zh-tw/dotnet/framework/data/transactions/implementing-an-explicit-transaction-using-committabletransaction
(四)建立 CommittableTransaction
1.
2.
3.
4.
5.

參考連結:https://dotblogs.com.tw/echo/2017/08/24/windows_msdtc_setting
參考連結:https://blog.darkthread.net/blog/category/MSDTC
23.3 DTC 分散式交易
1.
2.
3.
4.
5.

(二)
1.
2.
3.
4.
5.

(三)
1.
2.
3.
4.
5.

(四)
1.
2.
3.
4.
5.
**********************************************************************************************************
Module 33. TableAdapter
33.1	Update / AcceptChange
33.2	GetChanged / DataRowState - 資料列狀態
33.3	資料衝突

參考連結:https://docs.microsoft.com/zh-tw/visualstudio/data-tools/fill-datasets-by-using-tableadapters?view=vs-2019

目錄
1.TableAdapter 元件會根據程式指定的一或多個查詢or預存程式，將資料庫中的資料填入資料集→Fill()方法
2.Tableadapter 也可以在資料庫上執行作業→新增、刪除、更新
3.不過要注意的是，Tableadapter 是由 Visual Studio 設計工具產生的，我們無法直接撰寫程式碼使用TableAdapter
4.
5.
/*我自己的想法是TableAdapter將資料給DataSet後，但其他方法都是DataSet使用*/
datatable AcceptChange 參考連結:https://docs.microsoft.com/zh-tw/dotnet/api/system.data.datatable.acceptchanges?view=netframework-4.8

tableadapter.update :https://docs.microsoft.com/zh-tw/visualstudio/data-tools/update-data-by-using-a-tableadapter?view=vs-2019

33.1 Update / AcceptChange
1.TableAdapter 的更新功能取決於Tableadapter Wizard的主要查詢中有多少可用的資訊
2.
3.
4.
5.

(二)
1.
2.
3.
4.
5.

(三)
1.
2.
3.
4.
5.

(四)
1.
2.
3.
4.
5.

33.2
1.
2.
3.
4.
5.

(二)
1.
2.
3.
4.
5.

(三)
1.
2.
3.
4.
5.

(四)
1.
2.
3.
4.
5.

33.3
1.
2.
3.
4.
5.

(二)
1.
2.
3.
4.
5.

(三)
1.
2.
3.
4.
5.

(四)
1.
2.
3.
4.
5.

484要建立一個資料庫的資料表
架構簡單一點用來做測試
