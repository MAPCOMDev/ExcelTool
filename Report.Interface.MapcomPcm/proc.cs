using System;
using System.Data;
using System.Collections.Generic;
using System.Runtime.Serialization;

using PcmAxLib;

namespace Interface{
    /// <summary>
    /// MAPCOM PC-Mapping プロジェクトデータ処理拡張（処理）
    /// </summary>
    [Serializable]
    public class proc{
        #region 設定
        private string vLayerName = "屋外広告物";
        private int    vLayerType = 4;
        #endregion

        private main                    _o                 = null;
        private PcmAxLib.PcmAutoProject _AxProject         = null;
        private PcmAxLib.PcmAutoPcmDb   _AxDatabase        = null;
        private List<int>               _AxDatabaseTypeIdx = null;
        private int                     _AxDatabaseReadN   = 0;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="o">呼び出し元</param>
        public proc(main o){
            _o = o;
            _o.Log("処理：開始");

            string _vConfig = "";

            _vConfig = "既定";
            string _vLayerName = _o.ConfigRead("Layer");
            if(_vLayerName != ""){
                vLayerName = _vLayerName;
                _vConfig = "設定";

            }
            _o.Log("レイヤー名[" + _vLayerName + "]件（" + _vConfig + "値）");

            _vConfig = "既定";
            string _vLayerType = _o.ConfigRead("LayerType");
            if(_vLayerType != ""){
                try{
                    int _vLayerTypeN = int.Parse(_vLayerType);
                    if(_vLayerTypeN == 1
                    || _vLayerTypeN == 2
                    || _vLayerTypeN == 3
                    || _vLayerTypeN == 4
                    ){
                        vLayerType = _vLayerTypeN;
                        _vConfig   = "設定";
                    }
                }
                catch(Exception){
                }
            }
            _o.Log("レイヤータイプ[" + vLayerType + "]件（" + _vConfig + "値）");
        }

        /// <summary>
        /// 開始
        /// </summary>
        /// <param name="AxProject">PCMプロジェクト</param>
        public void Load(PcmAxLib.PcmAutoProject AxProject){
            _AxProject = AxProject;
            _o.Log("処理：プロジェクトロード");
        }

        /// <summary>
        /// 初期化
        /// </summary>
        public void Init(){
            _AxDatabase        = null;
            _AxDatabaseTypeIdx = null;
            _AxDatabaseReadN   = 0;
        }

        /// <summary>
        /// [定義] 属性情報：定義
        /// </summary>
        /// <returns>処理結果</returns>
        public bool DatabaseType(){
            bool ret = false;

            Init();
            if(_AxProject != null){
                if(_AxProject != null){
                    int nN     = -1;
                    int nLayer = _AxProject.GetNumOfLayer();
                    for(int n = 0; n < nLayer; n++){
                        PcmAxLib.PcmAutoPcmLayer oLayer = _AxProject.GetLayer(n, true);
                        if(oLayer != null){
                            if(oLayer.Title == vLayerName){
                                nN = n;
                                break;
                            }
                        }
                    }

                    if(nN >= 0){
                        PcmAxLib.PcmAutoPcmLayer oLayer = _AxProject.GetLayer(nN, true);
                        if(oLayer != null){
                            _AxDatabase = oLayer.GetDb(vLayerType);
                            if(_AxDatabase != null){
                                _o._DatabaseTypeInit();

                                _AxDatabaseTypeIdx = new List<int>();

                                int nDatabaseField = _AxDatabase.GetNumOfField();
                                for(int iDatabaseField = 0; iDatabaseField < nDatabaseField; iDatabaseField++){
                                    int    t    = _AxDatabase.FieldType[iDatabaseField];
                                    string name = _AxDatabase.FieldName[iDatabaseField];

                                    if(name == "@調査日時"
                                    || name == "@区分"
                                    || name == "@表示内容"
                                    || name == "@設置場所"
                                    || name == "@住所"
                                    || name == "@連絡先"
                                    || name == "座標系"
                                    || name == "X座標"
                                    || name == "Y座標"
                                    || name == "更新日時"
                                    || name == "UID"
                                    ){ 
                                        _o.LogMsg("処理：データベース定義：フォールド名[" + name + "]を含める");
                                    }
                                    else{
                                        _o.LogMsg("処理：データベース定義：フォールド名[" + name + "]を除外");
                                        continue;
                                    }

                                    _o._DatabaseTypeSet(name, t);
                                    _AxDatabaseTypeIdx.Add(iDatabaseField);

                                    ret = true;
                                }
                            }
                            else{
                                _o.LogMsg("データベースへアクセスできません");
                            }
                        }
                        else{
                            _o.LogMsg("レイヤー名[" + vLayerName + "]（" + nN + "）へアクセスできません");
                        }
                    }
                    else{
                        _o.LogMsg("レイヤー名[" + vLayerName + "]が存在しません");
                    }
                }
            }
            else{
                _o.LogMsg("プロジェクトが存在しません");
            }

            _o.LogSave();

            return ret;
        }

        /// <summary>
        /// [定義] 属性情報：データ：行読込
        /// </summary>
        /// <returns>処理結果</returns>
        public bool DatabaseRead(){
            bool ret = false;

            DataTable Type = _o._DatabaseTypeGet();

            if(_AxDatabase != null
            && Type        != null                    
            ){
                if(_AxDatabaseReadN < _AxDatabase.GetNumOfRec()){
                    _o.Log("処理：データベース読込[" + (_AxDatabaseReadN + 1).ToString() + "]行目");
                    _o._DatabaseReadInit();
                    for(int n = 0; n < Type.Rows.Count; n++){
                        string v = _AxDatabase.GetCell(_AxDatabaseReadN, _AxDatabaseTypeIdx[n], true);
                        if(_o._DatabaseReadSet(n, v)){
                            ret = true;
                        }
                        else{
                            ret = false;
                        }
                    }
                    _AxDatabaseReadN++;
                }
            }
            return ret;
        }
  

        /// <summary>
        /// [定義] 解放
        /// </summary>
        public void Dispose(){
            Init();
            _o.Log("処理：終了");
        }
    }
}