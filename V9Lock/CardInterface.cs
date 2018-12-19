using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;

/// <summary>Интерфейс с объявлениями пользовательских методов и свойств компоненты</summary>
/// 
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ICardInterface
{
    /// <summary>Пример реализации свойства компоненты</summary>
    string Prop { get; set; }

    /// <summary>Пример реализации метода компоненты</summary>
    int Go(int Param1, int Param2);

    int Buzz();

    string GetDllVersion();

    string ReadCardData();

    string GetGuestCardinfo(string coid);

    string GetCoID();

    string WriteGuestCard(string coid, string lockcode, int cardnum, int daifo, int open, DateTime date);


}

namespace AddIn
{
    /// <summary>Класс, реализующий пользовательские методы компоненты</summary>
    [ProgId("AddIn.CardInterface")]
    [ClassInterface(ClassInterfaceType.None)]
    public class CardInterface : AddIn, ICardInterface
    {
        //Инициализация
        [DllImport("V9RF.dll", EntryPoint = "Buzzer")]
        public static extern int Buzzer(int sc);

        //Прочитать номер версии DLL
        [DllImport("V9RF.dll", EntryPoint = "GetDLLVersionA")]
        public static extern int GetDLLVersionA(byte[] aType);

        //Прочитать логотип отеля у эмитента карты
        [DllImport("V9RF.dll", EntryPoint = "Getcoid")]
        public static extern int Getcoid(byte[] dlscoid);

        //Данные считывания карты
        [DllImport("V9RF.dll", EntryPoint = "ReadCard")]
        public static extern int ReadCard(byte[] carddata);

        //Выходная карта
        [DllImport("V9RF.dll", EntryPoint = "CardEraseA")]
        public static extern int CardEraseA(int coid, byte[] carddata);

        //Гостевая карта
        [DllImport("V9RF.dll", EntryPoint = "WriteGuestCardA")]
        public static extern int WriteGuestCardA(int dlscoid, byte cardno, byte dai, byte llock, char[] EDate, char[] RoomNo, byte[] cardhexstr);


        //Получить информацию о гостевой карте
        [DllImport("V9RF.dll", EntryPoint = "GetGuestCardinfoA")]
        public static extern int GetGuestCardinfoA(int dlscoid, byte[] carddata, byte[] lockno);



        byte[] carddata = new byte[128];
        /// <summary>Хранилище для значения свойства Prop</summary>
        private string prop;
        

        /// <summary>Реализация свойства Prop</summary>
        public string Prop
        {
            get
            {
                return prop;
            }
            set
            {
                prop = value;
            }

        }

        /// <summary>Функция для проверки доступности компоненты</summary>
        public int Go(int Param1, int Param2)
        {
            try
            {
                return (Param1 / Param2);
            }
            catch (Exception e)
            {
                asyncEvent.ExternalEvent("AddIn", "error", e.ToString());
                return 0;
            }
        }
        public int Buzz()
        {
            Buzzer(50);
            return 0;
        }
        public string GetDllVersion()
        {
            byte[] types = new byte[20];
            string strp = "";
            int i, st;
            st = GetDLLVersionA(types);
            if (st == 0)
            {
                for (i = 0; i < 20; i++)
                {
                    strp = strp + ((char)types[i]).ToString();
                }
                return "Версия DLL：" + strp;
            }
            else return "Getdll Fails!";
        }

        public string ReadCardData()
        {
            int i, st;
            string datastr = "";
            byte[] cardbuf = new byte[128];
            st = ReadCard(cardbuf);
            Thread.Sleep(400);//Рекомендуется задерживать 500 миллисекунд, ожидая ответа оборудования
            if (st == 0)
            {
                Buzzer(50);
                for (i = 0; i < 38; i++)
                    datastr = datastr + ((char)cardbuf[i]).ToString();
                return datastr;

               
            }
            else
            {
                return "Не удалось прочитать карту：" + st.ToString();
            }
        }

        public string GetGuestCardinfo(string coid)
        {
            byte[] lockno = new byte[50];

            int i, st;
            string locknostr = "", datastr = "", btime = "", etime = "", ulock = "", cardno = "";
            int dlscoid;
            dlscoid = Convert.ToInt32(coid);
            st = GetGuestCardinfoA(dlscoid, carddata, lockno);
            for (i = 0; i < 40; i++)
                datastr = datastr + ((char)carddata[i]).ToString();
            
            if (st == 0)
            {
                for (i = 0; i < 6; i++)
                    locknostr = locknostr + ((char)lockno[i]).ToString();
                for (i = 6; i < 18; i++)
                    btime = btime + ((char)lockno[i]).ToString();
                for (i = 18; i < 30; i++)
                    etime = etime + ((char)lockno[i]).ToString();
                for (i = 32; i < 40; i++)
                    cardno = cardno + ((char)lockno[i]).ToString();
                ulock = ulock + ((char)lockno[30]).ToString();

                return "Номер карты:" + cardno + "\n" + "Номер блокировки:" + locknostr + "\n" + "Время выпуска карты:" + btime + "\n" + "Срок годности:" + etime + "\n" + "Открывать ли замок:" + ulock;

            }
            else if (st == 1)
            {
                return "Не удалось подключиться к эмитенту карты, вернуть значение：" + st.ToString();

            }
            else if (st == -2)
            {
                return "Нет действительной карты, возвращаемое значение：" + st.ToString();

            }
            else if (st == -3)
            {
                return "Не-отельная карта, логотип отеля не соответствует, возвращаемая стоимость：" + st.ToString();

            }
            else if (st == -4)
            {
                return " Пустая карточка или карта, которая была вышла из системы, возвращает значение：" + st.ToString();

            }
            else
            {
                return "Неизвестное возвращаемое значение：" + st.ToString();

            }

        }
        public string GetCoID()
        {
            int i, k;
            byte[] Dcoid = new byte[20];
            string coid = "";
            k = Getcoid(Dcoid);
            if (k == 0)
            {
                for (i = 0; i < 6; i++)
                    coid = coid + ((char)Dcoid[i]).ToString();
                //textBox5.Text = coid;
                i = Convert.ToInt32(coid.Substring(0, 2), 16) * 65536 + Convert.ToInt32(coid.Substring(2, 4), 16) % 16383;
                
                return i.ToString();

            }
            return "Не удалось получить идентификатор гостиницы " + k.ToString();
        }
        public string WriteGuestCard(string coid, string lockcode,int cardnum, int daifo,int open,DateTime date)
        {
            int i, st;
            int dlscoid;
            byte cardno;
            byte dai;
            byte llock;
            string datastr = "";
            string lockstr, EDatestr;
            byte[] cardbuf = new byte[128];

            char[] lockno = new char[6];
            char[] EDate = new char[10];

            lockstr = lockcode;//Номер блокировки
            
            for (i = 0; i < 6; i++)
                lockno[i] = Convert.ToChar(lockstr.Substring(i, 1));
            EDatestr = date.ToString("yyMMddHHmm");//dateTimePicker1.Value.ToString("yyMMdd") + dateTimePicker2.Value.ToString("HHmm");//Эффективное время
            for (i = 0; i < 10; i++)
                EDate[i] = Convert.ToChar(EDatestr.Substring(i, 1));

            dlscoid = Convert.ToInt32(coid);//Логотип гостиницы
            cardno = Convert.ToByte(cardnum);  //Номер карты 0..15
            dai = Convert.ToByte(daifo);   //Ширина передней панели карты 0..255.
            llock = Convert.ToByte(open);  //Открыть антиблокировочный знак
            

            st = WriteGuestCardA(dlscoid, cardno, dai, llock, EDate, lockno, cardbuf);
            Thread.Sleep(400);//Рекомендуется задержать 400 миллисекунд, ожидая аппаратного ответа
            if (st == 0)
            {
                Buzzer(50);
                for (i = 0; i < 32; i++)
                    datastr = datastr + ((char)carddata[i]).ToString();
                return datastr;

                
            }
            else
            {
                return "Ошибка карты гостя：" + st.ToString();
            }
        }

    }
}
