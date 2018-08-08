using System;
using System.IO;
using System.Data;
using System.Reflection;
using System.Runtime.Serialization;

using PcmAxLib;

namespace Interface{
    /// <summary>
    /// MAPCOM PC-Mapping プロジェクトデータ処理拡張
    /// </summary>
    [Serializable]
    public class main : Interface.MapcomPcm{
        private string                  _vDLL               = "";
        private string                  _vPath              = "";

        private Interface.Log           _oLog               = null;
        private Interface.Config        _oConfig            = null;
        private proc                    _oProc              = null;

        private PcmAxLib.PcmAutoProject _AxProject          = null;
        private DataTable               _AxDatabaseType     = null;
        private string []               _AxDatabaseDataLine = null;

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
        /// [定義] プロジェクト
        /// </summary>
        /// <param name="AxProject"></param>
        public void Project(PcmAxLib.PcmAutoProject AxProject){        
            _AxProject = AxProject;
            if(_oProc != null){
                _oProc.Load(_AxProject);
            }
        }

        /// <summary>
        /// [定義] 属性情報：定義
        /// </summary>
        /// <returns>定義情報</returns>
        public DataTable DatabaseType(){
            DataTable ret = null;
            if(_oProc != null){
                if(_oProc.DatabaseType()){
                    ret = _AxDatabaseType;
                }
            }
            return ret;
        }


        /// <summary>
        /// [定義] 属性情報：データ：行読込
        /// </summary>
        /// <returns>行データ</returns>
        public string[] DatabaseRead(){
            _AxDatabaseDataLine = null;
            string [] ret = null;

            if(_oProc != null){
                if(_oProc.DatabaseRead()){
                    ret = _AxDatabaseDataLine;
                }
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

            _AxProject  = null;
        }
        #endregion

        #region 定義 for class proc
        /// <summary>
        /// 属性情報：定義：初期化
        /// </summary>
        public void _DatabaseTypeInit(){
            _AxDatabaseType = new DataTable();
            _AxDatabaseType.Columns.Add("フィールド名");
            _AxDatabaseType.Columns.Add("フィールド型");
        }

        /// <summary>
        /// 属性情報：定義：取得
        /// </summary>
        /// <returns></returns>
        public DataTable _DatabaseTypeGet(){
            return _AxDatabaseType;
        }

        /// <summary>
        /// 属性情報：定義：設定
        /// </summary>
        /// <param name="name">フィールド名</param>
        /// <param name="type">フィールド型番号</param>
        public void _DatabaseTypeSet(string name, int t){
            string type = "文字型";
            if     (t ==  0){ type = "文字型"             ; }
            else if(t ==  1){ type = "整数型"             ; }
            else if(t ==  2){ type = "整数型(先行ゼロ型) "; }
            else if(t ==  3){ type = "長整数型"           ; }
            else if(t ==  4){ type = "実数型"             ; }
            else if(t ==  5){ type = "BOOL型"             ; }
            else if(t ==  6){ type = "度分秒型"           ; }
            else if(t ==  8){ type = "DECIMAL型"          ; }
            else if(t ==  9){ type = "数値型"             ; }
            else if(t == 11){ type = "フォント型"         ; }
            else if(t == 12){ type = "カラー型"           ; }
            else if(t == 13){ type = "文字サイズ型"       ; }
            else if(t == 14){ type = "マスク型"           ; }
            else if(t == 15){ type = "文字整列型"         ; }
            else if(t == 16){ type = "回転角度型"         ; }
            else if(t == 17){ type = "サークル表示型"     ; }
            else if(t == 18){ type = "アーク矢印型"       ; }
            else if(t == 22){ type = "文字ボックス型"     ; }
            else if(t == 23){ type = "日付時刻型"         ; }
            else if(t == 24){ type = "縮尺制御型"         ; }
            else if(t == 25){ type = "注記属性型"         ; }
            else if(t == 26){ type = "チャート型"         ; }
            else if(t == 27){ type = "注記属性B型"        ; }
            else if(t == 30){ type = "BLOB型"             ; }
            else if(t == 31){ type = "データベース型"     ; }
            else if(t == 32){ type = "列挙型"             ; }
            _DatabaseTypeSet(name, type);
        }

        /// <summary>
        /// 属性情報：定義：設定
        /// </summary>
        /// <param name="name">フィールド名</param>
        /// <param name="type">フィールド型名</param>
        public void _DatabaseTypeSet(string name, string type){
            string[] r = new string[2];
            r[0] = name;
            r[1] = type;
            _AxDatabaseType.Rows.Add(r);
        }

        /// <summary>
        /// 属性情報：データ：行読込：初期化
        /// </summary>
        /// <returns></returns>
        public bool _DatabaseReadInit(){
            bool ret = false;
            _AxDatabaseDataLine = null;

            if(_AxDatabaseType != null){
                _AxDatabaseDataLine = new string[_AxDatabaseType.Rows.Count];
                ret = true;
            }
            return ret;
        }

        /// <summary>
        /// 属性情報：データ：行読込：設定
        /// </summary>
        /// <param name="n">列番号[ 0 オリジン]</param>
        /// <param name="v">値</param>
        /// <returns></returns>
        public bool _DatabaseReadSet(int n, string v){
            bool ret = false;
            if(_AxDatabaseDataLine != null){
                if(n < _AxDatabaseDataLine.Length){
                    _AxDatabaseDataLine[n] = v;
                    ret = true;
                }
            }
            return ret;
        }
        #endregion
    }
}