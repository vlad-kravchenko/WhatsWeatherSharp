using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace WhatsWeatherSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Объект основного класса для хранения ответа с сервера API
        WeatherClass myClass;
        private async void button1_Click(object sender, EventArgs e)
        {
            WebRequest request = null;
            OleDbConnection connection;
            string ansver = string.Empty;
            DataTable dtAccess = new DataTable();
            //Вход в функцию обработки клика по кнопке с учётом возможных исключений
            try
            {
                //Сначала подключаемся к БД (вдруг запрос сегодня уже посылался?)
                String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=WeatherDB.mdb";
                String queryString = string.Empty;
                OleDbDataAdapter adapter = null;
                connection = new OleDbConnection(connectionString);
                connection.Open();
                //Запрашиваем всё по конкретному введённому городу за текущую дату (в двух вариантах - по названию и ID)
                if (comboBox1.Text == "названию" && textBox1.Text != null && dtAccess.Rows.Count == 0)
                {
                    queryString = "SELECT * FROM WeatherStorage WHERE dateNow=Date() AND sityQuerry='" + textBox1.Text + "'";
                    adapter = new OleDbDataAdapter(queryString, connectionString);
                    adapter.Fill(dtAccess);
                    //Здесь же создаём заготовку запроса чтобы не прописывать условие ещё раз (на случай, если в БД записей нет)
                    request = WebRequest.Create("http://api.openweathermap.org/data/2.5/forecast?q=" + textBox1.Text + "&units=metric&lang=ru&APPID=0a090c428a060cac715bb45a87acb320");
                }
                //Вторая проверка по названию города - на случай, если сегодня этот же город запрашивался на другом языке
                //Таким образом, в БД в сутки по одному городу хранится НЕ БОЛЕЕ трёх записей - в русской/украинской/английской версиях
                if (comboBox1.Text == "названию" && textBox1.Text != null && dtAccess.Rows.Count == 0)
                {
                    queryString = "SELECT * FROM WeatherStorage WHERE dateNow=Date() AND sity='" + textBox1.Text + "'";
                    adapter = new OleDbDataAdapter(queryString, connectionString);
                    adapter.Fill(dtAccess);
                    request = WebRequest.Create("http://api.openweathermap.org/data/2.5/forecast?q=" + textBox1.Text + "&units=metric&lang=ru&APPID=0a090c428a060cac715bb45a87acb320");
                }
                else if (comboBox1.Text == "ID" && textBox1.Text != null && dtAccess.Rows.Count == 0)
                {
                    queryString = "SELECT * FROM WeatherStorage WHERE dateNow=Date() AND sityID=" + textBox1.Text + "";
                    adapter = new OleDbDataAdapter(queryString, connectionString);
                    adapter.Fill(dtAccess);
                    request = WebRequest.Create("http://api.openweathermap.org/data/2.5/forecast?id=" + textBox1.Text + "&units=metric&lang=ru&APPID=0a090c428a060cac715bb45a87acb320");
                }
                //Если результат запроса не содержит данных, работаем с API
                if (dtAccess.Rows.Count==0)
                {
                    //Отправка запроса и получение ответа
                    request.Method = "POST";
                    request.ContentType = "application/x-www-urlencoded";
                    WebResponse response = await request.GetResponseAsync();
                    Stream s = response.GetResponseStream();
                    StreamReader reader = new StreamReader(s);
                    //Запись ответа в строку
                    ansver = await reader.ReadToEndAsync();
                    //Закрываем соединение с API
                    response.Close();
                    linkLabel19.Text = "Загружено: с сервера";
                }
                else
                {
                    //Срабатывает если подобный запрос сегодня уже был - в таком случае просто грузим всё из БД
                    ansver = dtAccess.Rows[0].ItemArray[4].ToString();
                    linkLabel19.Text = "Загружено: из базы данных";
                }
                //Парсим строку JSON в объект класса
                myClass = JsonConvert.DeserializeObject<WeatherClass>(ansver);
                //Вызов функции начального заполнения формы (по ближайшему прогнозу)
                FillFields(0);
                //Все функции по возможности ставим после заполнения, т.к. они требуют времени, а юзер в это время как раз будет изучать данные
                //Очистка (на всякий случай) comboBox, содержащего периоды прогноза...
                comboBox2.Items.Clear();
                //... и его заполнение новыми данными
                for (int i = 0; i < myClass.list.Length; i++)
                {
                    comboBox2.Items.Add(myClass.list[i].dt_txt);
                }
                //Устанавливаем корректную версию браузера
                SetWebBrowserCompatiblityLevel();
                //Отображение нужного города на карте
                webBrowser1.Navigate("https://www.google.com.ua/maps/place/" + myClass.city.name);
                //Чистим БД, так как хранить нам нужно только запросы за сегодня
                queryString = "DELETE FROM WeatherStorage WHERE dateNow<Date()";
                OleDbCommand myCmd = new OleDbCommand(queryString, connection);
                myCmd.ExecuteReader();
                //Записываем JSON, текущую дату и название города в БД (только если их там ещё нет)
                if (dtAccess.Rows.Count == 0)
                {
                    queryString = "INSERT INTO WeatherStorage(sityQuerry,sity,sityID,dateNow,stringJSON) VALUES ('" + textBox1.Text + "','" + myClass.city.name + "',"+ myClass.city.id+ ",Date(), '" + ansver + "')";
                    myCmd = new OleDbCommand(queryString, connection);
                    myCmd.ExecuteReader();
                    connection.Close();
                }
                //Очистка/обнуление данных
                connection.Close();
                request = null;
                ansver = string.Empty;
                dtAccess.Clear();
                connectionString = string.Empty;
                queryString = string.Empty;
                adapter = null;
            }
            catch (Exception ex)
            {
                //В случае ошибки выводим её текст
                MessageBox.Show(ex.Message.ToString(), "Error!");
            }
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //При выборе даты и времени отображается соответствующая часть прогноза
            int id_focus = comboBox2.SelectedIndex;
            FillFields(id_focus);
        }
        public void FillFields(int index)
        {
            //Вывод данных на форму
            //В блок "Подробно"-------------------------------------------------------
            linkLabel17.Text = myClass.city.name;
            linkLabel16.Text = myClass.city.id.ToString();
            linkLabel15.Text = myClass.city.country;
            linkLabel14.Text = myClass.list[index].main.temp.ToString("0.##");
            linkLabel13.Text = (myClass.list[index].main.pressure / 1.3332239).ToString("0.");
            linkLabel12.Text = myClass.list[index].main.humidity.ToString("0.");
            linkLabel11.Text = myClass.list[index].wind.speed.ToString("0.##");
            linkLabel18.Text = myClass.list[index].weather[0].description;
            pictureBox1.Image = Image.FromFile("icons/" + myClass.list[index].weather[0].icon + ".png");
            //Конвертируем направление ветра из градусов в привычные понятия
            int wind = (int)myClass.list[index].wind.deg;
            string wind_text = (wind < 30) ? "северный" : (wind < 60) ? "северо-восточный" : (wind < 120) ? "восточный" : (wind < 150) ? "юго-восточный" : 
                (wind < 210) ? "южный" : (wind < 240) ? "юго-западный" : (wind < 300) ? "западный" : (wind < 330) ? "северо-западный" : "северный";
            linkLabel10.Text = wind_text;
            //В блок "Кратко"---------------------------------------------------------
            //Находим первое значение по времени = 12:00 - это первый день прогноза
            int first_day = 0;
            for (int i = 0;i < myClass.list.Length; i++)
            {
                if (myClass.list[i].dt_txt.Contains("12:00"))
                {
                    first_day = i;
                    break;
                }
            }
            //Пробегаемся циклом, заполняя кратние сводки на каждый день (с учётом шага прогноза в 3 часа, разница по суткам будет +8 шагов)
            //Пробегаемся по groupBox-ам (которых всего 5 шт)
            for (int i = 4; i > -1; i --)
            {
                tabControl1.TabPages[0].Controls[i].Text = myClass.list[first_day].dt_txt.Remove(11);
                tabControl1.TabPages[0].Controls[i].Controls[0].Text = myClass.list[first_day].main.temp.ToString("0.##");
                tabControl1.TabPages[0].Controls[i].Controls[1].Text = myClass.list[first_day].wind.speed.ToString("0.##");
                int wind_last = (int)myClass.list[first_day].wind.deg;
                string wind_text_last = (wind_last < 30) ? "северный" : (wind_last < 60) ? "северо-восточный" : (wind_last < 120) ? "восточный" :
                    (wind_last < 150) ? "юго-восточный" : (wind_last < 210) ? "южный" : (wind_last < 240) ? "юго-западный" : (wind_last < 300) ? "западный" :
                    (wind_last < 330) ? "северо-западный" : "северный";
                tabControl1.TabPages[0].Controls[i].Controls[2].Text = wind_text_last;
                //Обратиться к PictureBox так же, как и к label нельзя, поэтому так:
                foreach (PictureBox pb in tabControl1.TabPages[0].Controls[i].Controls.OfType<PictureBox>())
                {
                    pb.Image = Image.FromFile("icons/" + myClass.list[first_day].weather[0].icon + ".png");
                }
                //Идём вперёд по суткам на 8 шагов (8*3=24)
                first_day += 8;
                if (first_day > myClass.list.Count()) break;
            }
        }
        //Ускоряем работу с программой
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
            }
        }
        //Начальная инициализация
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = comboBox1.Items[0];
            textBox1.Text = "Мелитополь";
            button1_Click(sender, e);
        }
        //Кратенький Help&About
        private void button2_Click(object sender, EventArgs e)
        {
            String about = @"Простой погодный информер, основанный на API сайта OpenWeatherMap. 
Для работы достаточно выбрать тип запроса и ввести название/ID города. 
Присутствует возможность просматривать погоду в двух вариантах - кратко, сразу на 5 дней и подробно, с шагом каждые 3 часа.

Разработчик: Sirius, 2017";
            String head = "О программе";
            MessageBox.Show(about, head, MessageBoxButtons.OK);
        }
        //Функции SetWebBrowserCompatiblityLevel, WriteCompatiblityLevel и GetBrowserVersion отвечают за корректную работу браузера в соответствии
        //с установленной на конкретной машине версией IE, так как по умолчанию используется IE7, который воспринимается Google как устаревший
        private static void SetWebBrowserCompatiblityLevel()
        {
            string appName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);
            int lvl = 1000 * GetBrowserVersion();
            bool fixVShost = File.Exists(Path.ChangeExtension(Application.ExecutablePath, ".vshost.exe"));
            WriteCompatiblityLevel("HKEY_LOCAL_MACHINE", appName + ".exe", lvl);
            if (fixVShost) WriteCompatiblityLevel("HKEY_LOCAL_MACHINE", appName + ".vshost.exe", lvl);
            WriteCompatiblityLevel("HKEY_CURRENT_USER", appName + ".exe", lvl);
            if (fixVShost) WriteCompatiblityLevel("HKEY_CURRENT_USER", appName + ".vshost.exe", lvl);
        }
        private static void WriteCompatiblityLevel(string root, string appName, int lvl)
        {
            try
            {
                Microsoft.Win32.Registry.SetValue(root + @"\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", appName, lvl);
            }
            catch (Exception) { }
        }
        public static int GetBrowserVersion()
        {
            string strKeyPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer";
            string[] ls = new string[] { "svcVersion", "svcUpdateVersion", "Version", "W2kVersion" };
            int maxVer = 0;
            for (int i = 0; i < ls.Length; ++i)
            {
                object objVal = Microsoft.Win32.Registry.GetValue(strKeyPath, ls[i], "0");
                string strVal = Convert.ToString(objVal);
                if (strVal != null)
                {
                    int iPos = strVal.IndexOf('.');
                    if (iPos > 0)
                        strVal = strVal.Substring(0, iPos);

                    int res = 0;
                    if (int.TryParse(strVal, out res))
                        maxVer = Math.Max(maxVer, res);
                }
            }
            return maxVer;
        }
    }
}
