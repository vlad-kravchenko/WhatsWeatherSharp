namespace WhatsWeatherSharp
{
    //Данный класс отображает структуру JSON-ответа и содержит основные данные для отображения пользователю
    class WeatherClass
    {
        public List[] list { get; set; }
        public City city { get; set; }
        public class List
        {
            public string dt_txt { get; set; }
            public Main main { get; set; }
            public Wind wind { get; set; }
            public Weather[] weather { get; set; }
            public class Main
            {
                public double temp { get; set; }
                public double pressure { get; set; }
                public int humidity { get; set; }
            }
            public class Wind
            {
                public double speed { get; set; }
                public double deg { get; set; }
            }
            public class Weather
            {
                public string description { get; set; }
                public string icon { get; set; }
            }
        }
        public class City
        {
            public string name { get; set; }
            public int id { get; set; }
            public string country { get; set; }
        }
    }
}