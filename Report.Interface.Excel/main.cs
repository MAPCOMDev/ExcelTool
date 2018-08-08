using System;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;

namespace Interface{
    /// <summary>
    /// Excel処理拡張
    /// </summary>
    [Serializable]
    public class main : Interface.Excel{
        private string           _vDLL    = "";
        private string           _vPath   = "";
        
        private Interface.Log    _oLog    = null;
        private Interface.Config _oConfig = null;
        private proc             _oProc   = null;

        private Hashtable        _vEnt    = null;
        private DataTable        _vData   = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public main(){
            _vDLL    = Assembly.GetExecutingAssembly().Location;            
            _vPath   = Path.GetDirectoryName(_vDLL);

            _oLog    = new Log   (_vDLL);
            _oConfig = new Config(_vDLL);
            _oProc   = new proc  (this );
        }

        #region ログ
        /// <summary>
        /// ログ
        /// </summary>
        /// <param name="msg">メッセージ</param>
        public void Log    (string msg){ _oLog.SaveMsg(msg); }

        /// <summary>
        /// ログ：記録
        /// </summary>
        /// <param name="msg">メッセージ</param>
        public void LogMsg (string msg){ _oLog.Msg(msg);     }

        /// <summary>
        /// ログ：保存
        /// </summary>
        public void LogSave(          ){ _oLog.Save();       }
        #endregion

        #region 設定
        /// <summary>
        /// 設定：読込
        /// </summary>
        /// <param name="key">キー</param>
        /// <returns></returns>
        public string ConfigRead (string key          ){ return _oConfig.Read (key   ); }

        /// <summary>
        /// 設定：書込
        /// </summary>
        /// <param name="key">キー</param>
        /// <param name="v">値</param>
        /// <returns></returns>
        public void   ConfigWrite(string key, string v){        _oConfig.Write(key, v); }
        #endregion

        #region 定義
        /// <summary>
        /// [定義] 環境設定
        /// </summary>
        /// <param name="v">環境値</param>
        /// <param name="d">データ</param>
        public void Env(Hashtable v, DataTable d){
            _vEnt  = v;
            _vData = d;
        }

        /// <summary>
        /// [定義] [帳票形式] ページデータ：データ数
        /// </summary>
        /// <returns>ページデータ数</returns>
        public int PageData(){
            int ret = 0;
            if(_oProc != null){
                ret = _oProc.PageData();
            }
            return ret;
        }

        /// <summary>
        /// [定義] [帳票形式] ページデータ：セル：オフセット値
        /// </summary>
        /// <param name="cell">セル名</param>
        /// <param name="page">ページ[1～]</param>
        /// <param name="n">ページ内のデータ数[0～]</param>
        /// <param name="row">行移動量</param>
        /// <param name="col">列移動量</param>
        /// <returns></returns>
        public bool PageDataCellOffset(string cell, int page, int n, out int row, out int col){
            bool ret = false;
            row = 0;
            col = 0;
            if(_oProc != null){
                ret = _oProc.PageDataCellOffset(cell, page, n, out row, out col);
            }
            return ret;
        }

        /// <summary>
        /// [定義] データ
        /// </summary>
        /// <param name="page">ページ[1～]</param>
        /// <param name="n">ページ内のデータ数[0～]</param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <returns>
        /// Hashtable[cell]  セル名
        /// Hashtable[value] 値
        /// Hashtable[style] セル書式
        /// </returns>
        public List<Hashtable> Data(int page, int n, int row, int col){
            List<Hashtable> ret = null;
            if(_oProc != null){
                ret = _oProc.Data(page, n, row, col);
            }
            return ret;
        }

        /// <summary>
        /// [定義] データ：変換
        /// </summary>
        /// <param name="cell">セル名</param>
        /// <param name="v">値</param>
        /// <param name="style">セル書式</param>
        /// <param name="page">ページ[1～]</param>
        /// <param name="n">ページ内のデータ数[0～]</param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <returns>
        /// Hashtable[value] 値
        /// Hashtable[style] セル書式
        /// </returns>
        public Hashtable DataConvert(string cell, string v, Hashtable style, int page, int n, int row, int col){
            Hashtable ret = null;
            if(_oProc != null){
                ret = _oProc.DataConvert(cell, v, style, page, n, row, col);
            }
            return ret;
        }

        /// <summary>
        /// [定義] 解放
        /// </summary>
        public void Dispose(){
            if(_oProc != null){
                _oProc.Dispose();
            }
        }
        #endregion

        #region 定義 for class proc
        /// <summary>
        /// 環境設定：取得
        /// </summary>
        /// <param name="v">キー</param>
        /// <returns>値</returns>
        /// <remarks>
        /// ・呼び出し元 EXE
        /// [Exe]
        /// 
        /// ・呼び出し元 EXEパス
        /// [ExePath]
        /// 
        /// ・データ名
        /// [Data]
        /// 
        /// ・データパス
        /// [DataPath]
        /// 
        /// ・保存場所
        /// SaveType == 0    : フォルダ
        /// SaveType == 1, 2 : ファイル
        /// [Save]
        /// 
        /// ・保存タイプ
        /// 0 : 別々のファイルに出力
        /// 1 : １ファイルでシートに分けて出力
        /// 2 : １ファイルで１シートに範囲を指定して繰り返し出力
        /// [SaveType]
        /// 
        /// ・保存範囲（セル名、行列番号）
        /// SaveType ==   2 : 左上[A],右下[B]のセル名
        /// [SaveRangeA]
        /// [SaveRangeB]
        /// 
        /// ・保存範囲（左上）：行番号
        /// [SaveRangeA_Row]
        /// 
        /// ・保存範囲（左上）：列番号
        /// [SaveRangeA_Col]
        /// 
        /// ・保存範囲（右下）：行番号
        /// [SaveRangeB_Row]
        /// 
        /// ・保存範囲（右下）：列番号
        /// [SaveRangeB_Col]
        /// </remarks>
        public string _Env(string key){
            string ret = "";
            if(_vEnt != null){
                if(_vEnt.Contains(key)){
                    ret = (string) _vEnt[key];
                }
            }
            return ret;
        }
        #endregion
    }
}