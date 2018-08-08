using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Interface{
    /// <summary>
    /// Excel処理拡張（処理）
    /// </summary>
    [Serializable]
    public class proc{
        #region 設定
        private int PageDataN = 10;
        #endregion

        private main _o = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="o">呼び出し元</param>
        public proc(main o){
            _o = o;
            _o.Log("処理：開始");
        }

        /// <summary>
        /// [定義] [帳票形式] ページデータ：データ数
        /// </summary>
        /// <returns>ページデータ数</returns>
        public int PageData(){
            string _vConfig = "既定";
            string _PageDataN = _o.ConfigRead("PageDataN");
            if(_PageDataN != ""){
                try{
                    PageDataN = int.Parse(_PageDataN);
                    _vConfig = "設定";                    
                }
                catch(Exception){
                }
            }
            _o.Log("１ページのデータ数[" + _PageDataN + "]件（" + _vConfig + "値）");

            return PageDataN;
        }

        /// <summary>
        /// [定義] [帳票形式] ページデータ：セル：オフセット値
        /// </summary>
        /// <param name="cell">セル名</param>
        /// <param name="page">ページ[1～]</param>
        /// <param name="n">ページ内のデータ数[0～]</param>
        /// <param name="row">行移動量</param>
        /// <param name="col">列移動量</param>
        /// <returns>処理結果</returns>
        public bool PageDataCellOffset(string cell, int page, int n, out int row, out int col){
            bool ret = false;
            row = 0;
            col = 0;

            if(n > 0){
                cell = cell.ToLower();
                if(cell == "c3"
                || cell == "c4"
                || cell == "c5"
                || cell == "c6"
                || cell == "c7"
                || cell == "c8"
                || cell == "c9"
                || cell == "c10"
                || cell == "c11"
                || cell == "c12"
                || cell == "c16"
                ){
                    col = n;
                    _o.Log("セルオフセット：列方向＋[" + col.ToString() + "]");
                    ret = true;
                }
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

            if(n == 0){
                ret = new List<Hashtable>();

                Hashtable d = new Hashtable();
                string    dCell  = "B1";
                string    dValue = "DLL内部生成" + page.ToString();
                Hashtable dStyle = new Hashtable();
                dStyle["FontColor"] = "ff0000ff";

                d["cell" ] = dCell;
                d["value"] = dValue;
                d["style"] = dStyle;

                ret.Add(d);
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

            cell = cell.ToLower();
            if(cell == "c3"){
                string    dValue = v;
                Hashtable dStyle = new Hashtable();
                dStyle["FontColor"] = "ffff0000";

                ret = new Hashtable();
                ret["value"] = dValue;
                ret["style"] = dStyle;
            }

            return ret;
        }

        /// <summary>
        /// [定義] 解放
        /// </summary>
        public void Dispose(){
            _o.Log("処理：終了");
        }
    }
}