using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KR_DailyReport
{
    //********************************************************************************
    // MELSECNET/Hの通信モジュール
    //********************************************************************************
    public class NET10Control
    {
        // MELSECNET/H関連の定数
        const short CHANNEL = 51;       // 通信回線のチャネルNo.(MELSECNET/H（1 枚目）)
        const short MODE = -1;          // ダミー（-1固定）
        const int NETNO = 0x00;         // ネットワークNo.（自局）
        const int STNO = 0xFF;          // 局番（自局）
        const int DEVTP = 24;           // デバイスタイプ（リンクレジスタ(LW)）
        const int DEVNO = 0;            // 先頭デバイスNo.
        const int DTSIZE = 0x2;         // 書き込みバイトサイズ（1WORD）

        // 変数
        private bool m_Open;            // true : 通信中

        private int m_ChPath;           // オープンされた回線のパスのポインタ

        private short[] m_Buff;         // 送受信バッファ（指定されたアドレスのデータのみ格納）

        //********************************************************************************
        // コンストラクタ
        //********************************************************************************
        public NET10Control()
        {
            m_Open = false;
        }

        //********************************************************************************
        // デストラクタ
        //********************************************************************************
        ~NET10Control()
        {
            // 通信停止
            EndComm();
        }

        //********************************************************************************
        // NET10の通信開始
        //********************************************************************************
        public bool StartComm()
        {
            if (m_Open == true)
                return true;
            
            // 通信回線のオープン
            short ret = MDFUNC32.mdOpen( CHANNEL, MODE, ref m_ChPath );
            if ( (ret != 0) && (ret != 66))    // OPEN 済みエラーも正常終了とみなす
                return false;

            m_Open = true;

            return true;
        }

        //********************************************************************************
        // NET10の通信停止
        //********************************************************************************
        public bool EndComm()
        {
            if (m_Open == false)
                return true;

            // 通信回線のクローズ
            short ret = MDFUNC32.mdClose(m_ChPath);
            if (ret != 0)
                return false;

            m_Open = false;

            return true;
        }

        /// <summary>
        /// NET10の指定アドレスからデータを受信する
        /// </summary>
        /// <param name="NET10Info"></param>
        /// <returns>受信取得値</returns>
        public short DataReceive(NET10_SendAddressInfo NET10Info)
        {
            short size = DTSIZE;   // 受信データサイズ（WORDをBYTEにするため2倍）

            m_Buff = new short[1];

            // データを取得する
            int ret = MDFUNC32.mdReceive(m_ChPath, STNO, DEVTP, NET10Info.Addr[0], ref size, ref m_Buff[0]);

            return m_Buff[0];
        }

        /// <summary>
        /// NET10の指定アドレスへデータを送信する
        /// </summary>
        /// <param name="NET10Info"></param>
        /// <returns>送信結果</returns>
        public short DataSend(NET10_SendAddressInfo NET10Info)
        {
            short size = DTSIZE;   // 送信データサイズ（WORDをBYTEにするため2倍）

            m_Buff = new short[1];

            // 書き込みデータを格納
            m_Buff[0] = NET10Info.Data[0];
            short ret = MDFUNC32.mdSend(m_ChPath, STNO, DEVTP, NET10Info.Addr[0], ref size, ref m_Buff[0]);

            return ret;
        }

    }
}
