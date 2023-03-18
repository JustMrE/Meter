using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public class ColorsData
    {
        public HashSet<Color> allColors {get; set; }
        [JsonIgnore]
        public static List<Color> allColorsList { get; set; }
        public Dictionary<string, Color> subColors {get; set; }
        public Dictionary<string, Color> main {get; set; }
        public Dictionary<string, Color> mainTitle {get; set; }
        public Dictionary<string, Color> mainSubtitle {get; set; }
        public Dictionary<string, Color> extraTitle {get; set; }
        public Dictionary<string, Color> extraSubtitle {get; set; }
        public Color sumColor { get; set; }

        public static ColorsData instatnce;


        public static Dictionary<string, Color> colors = new Dictionary<string, Color>();
        public static Dictionary <string, Color> colorsForSettings = new Dictionary<string, Color>();
        public static Dictionary <string, Color> oldColorsForSettings;

        public ColorsData()
        {
            if (instatnce == null) instatnce = this;
        }

        public void CreateAllColors()
        {
            allColorsList = new List<Color>();
            allColorsList.AddRange(subColors.Values.ToArray());
            allColorsList.AddRange(main.Values.ToArray());
            allColorsList.AddRange(mainTitle.Values.ToArray());
            allColorsList.AddRange(mainSubtitle.Values.ToArray());
            allColorsList.AddRange(extraTitle.Values.ToArray());
            allColorsList.AddRange(extraSubtitle.Values.ToArray());
            allColorsList.Add(sumColor);
        }

        public static void LoadStandartColors()
        {
            SaveLoader.LoadStandartColors();
            UpdateColors();
        }
        
        public void GenerateClors()
        {
            allColors = new HashSet<Color>();
            subColors = new Dictionary<string, Color>();
            main = new Dictionary<string, Color>();
            mainTitle = new Dictionary<string, Color>();
            mainSubtitle = new Dictionary<string, Color>();
            extraTitle = new Dictionary<string, Color>();
            extraSubtitle = new Dictionary<string, Color>();

            Color color = ColorTranslator.FromHtml("#FF5050");
            subColors.Add("colorUp1", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#F0F000");
            subColors.Add("colorUp2", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#00B050");
            subColors.Add("colorUp3", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#8EA9DB");
            main.Add("subject", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#D9E1F2");
            main.Add("прием", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#D9E1F2");
            main.Add("отдача", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#D7E1F2");
            main.Add("сальдо", color);
            allColors.Add(color);

            color = ColorTranslator.FromHtml("#D7E1F2");
            main.Add("план", color);
            allColors.Add(color);

            AddColor("аскуэ", ColorTranslator.FromHtml("#66FF66"));
            AddColor("формула", ColorTranslator.FromHtml("#FF6600"));
            AddColor("ручное", ColorTranslator.FromHtml("#3399FF"));
            AddColor("оперативное", ColorTranslator.FromHtml("#FF66FF"));
            AddColor("по плану", ColorTranslator.FromHtml("#00D0A8"));
            AddColor("по счетчику", ColorTranslator.FromHtml("#FFD966"));
            AddColor("счетчик", ColorTranslator.FromHtml("#FFC819"));
            AddColor("корректировка факт", ColorTranslator.FromHtml("#2F75B5"));
            AddColor("утвержденный", ColorTranslator.FromHtml("#94FF29"));
            AddColor("корректировка", ColorTranslator.FromHtml("#71E200"));
            AddColor("заявка", ColorTranslator.FromHtml("#AAE600"));
        }

        void AddColor(string name, Color color)
        {
            Color newColor = color;
            mainTitle.Add(name, newColor);
            allColors.Add(newColor);

            newColor = ChangeColorBrightness(color, .15f);
            mainSubtitle.Add(name, newColor);
            allColors.Add(newColor);
            byte R = color.R;
            R = R < 254 ? (byte)(R + 1) : (byte)(R - 1);
            newColor = Color.FromArgb(R, color.G, color.B);
            extraTitle.Add(name.ToUpper(), newColor);
            allColors.Add(newColor);

            newColor = ChangeColorBrightness(newColor, .15f);
            extraSubtitle.Add(name.ToUpper(), newColor);
            allColors.Add(newColor);
        }
        public void ChangeColor(string name, Color newColor, Color oldColor)
        {
            if (mainTitle.ContainsKey(name))
            {
                mainTitle[name] = newColor;
                newColor = ChangeColorBrightness(newColor, .15f);
                mainSubtitle[name] = newColor;
            }
            else if (extraTitle.ContainsKey(name))
            {
                extraTitle[name] = newColor;

                allColorsList.Remove(oldColor);
                allColorsList.Add(newColor);

                oldColor = ChangeColorBrightness(oldColor, .15f);
                newColor = ChangeColorBrightness(newColor, .15f);

                extraSubtitle[name] = newColor;
                allColorsList.Remove(oldColor);
                allColorsList.Add(newColor);
            }

            
            //allColors.Add(newColor);
            //byte R = color.R;
            //R = R < 254 ? (byte)(R + 1) : (byte)(R - 1);
            //newColor = Color.FromArgb(R, color.G, color.B);
            //extraTitle.Add(name.ToUpper(), newColor);
            //allColors.Add(newColor);

            //newColor = ChangeColorBrightness(newColor, .15f);
            //extraSubtitle.Add(name.ToUpper(), newColor);
            //allColors.Add(newColor);
        }

        public Color? GetColor(string name)
        {
            if (mainTitle.ContainsKey(name))
            {
                return mainTitle[name];
            }
            else if (extraTitle.ContainsKey(name))
            {
                return extraTitle[name];
            }
            return null;
        }

        public static Color ChangeColorBrightness(Color color, float correctionFactor)
        {
            float red = (float)color.R;
            float green = (float)color.G;
            float blue = (float)color.B;

            if (correctionFactor < 0)
            {
                correctionFactor = 1 + correctionFactor;
                red *= correctionFactor;
                green *= correctionFactor;
                blue *= correctionFactor;
            }
            else
            {
                red = (255 - red) * correctionFactor + red;
                green = (255 - green) * correctionFactor + green;
                blue = (255 - blue) * correctionFactor + blue;
            }

            return Color.FromArgb(color.A, (int)red, (int)green, (int)blue);
        }

        public bool IsColorFree(Color color)
        {
            byte R = color.R;
            R = R < 254 ? (byte)(R + 1) : (byte)(R - 1);
            Color newColor = Color.FromArgb(R, color.G, color.B);

            return !allColorsList.Contains(color) && !allColorsList.Contains(ChangeColorBrightness(color, .15f)) && !allColorsList.Contains(newColor);
        }

        /*public static void RefreshColors()
        {
            colors.Add("colorUp1", ColorTranslator.FromHtml("#FF5050"));
            colors.Add("colorUp2", ColorTranslator.FromHtml("#F0F000"));
            colors.Add("colorUp3", ColorTranslator.FromHtml("#00B050"));
            colors.Add("subject", ColorTranslator.FromHtml("#8EA9DB"));
            colors.Add("in", ColorTranslator.FromHtml("#D9E1F2"));
            colors.Add("out", ColorTranslator.FromHtml("#D9E1F2"));
            colors.Add("saldo", ColorTranslator.FromHtml("#D7E1F2"));
            colors.Add("plan", ColorTranslator.FromHtml("#D7E1F2"));

            colorsForSettings.Add("аскуэ", ColorTranslator.FromHtml("#99FF99"));
            colorsForSettings.Add("формула", ColorTranslator.FromHtml("#FF9966"));
            colorsForSettings.Add("ручное", ColorTranslator.FromHtml("#66CCFF"));
            colorsForSettings.Add("оперативное", ColorTranslator.FromHtml("#FF99FF"));
            colorsForSettings.Add("по плану", ColorTranslator.FromHtml("#00FFCC"));
            colorsForSettings.Add("по счетчику", ColorTranslator.FromHtml("#FFE699"));
            colorsForSettings.Add("счетчик", ColorTranslator.FromHtml("#FFD44B"));
            colorsForSettings.Add("корректировка факт", ColorTranslator.FromHtml("#4C91D0"));
            colorsForSettings.Add("утвержденный", ColorTranslator.FromHtml("#CCFF99"));
            colorsForSettings.Add("корректировка", ColorTranslator.FromHtml("#99FF33"));
            colorsForSettings.Add("заявка", ColorTranslator.FromHtml("#CCFF33")); 
        }*/

        public static void UpdateColors()
        {

            MessageBox.Show("Это займет некоторое время. Дождитесь уведомления о завершении!");
            var watch = System.Diagnostics.Stopwatch.StartNew();
            Main.instance.StopAll();
            Main.instance.references.UpdateAllColors1(false);
            Main.instance.heads.UpdateAllColors(false);
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Готово!\n" + (watch.ElapsedMilliseconds / 1000) + " ms");
        }

        public static Color GetRangeColor(Excel.Range rng)
        {
            return ColorTranslator.FromOle(Convert.ToInt32(rng.Interior.Color));
        }
        public static bool IsRangeColor(Excel.Range rng, string name)
        {
            Color c = GetRangeColor(rng);
            return IsRangeColor(c, name);
        }
        public static bool IsRangeColor(Color c, string name) 
        {
            Dictionary<string, Color> d = null;
            if (instatnce.subColors.ContainsKey(name))
            {
                d = instatnce.subColors;
            }
            else if (instatnce.main.ContainsKey(name))
            {
                d = instatnce.main;
            }
            else if (instatnce.mainTitle.ContainsKey(name))
            {
                d = instatnce.mainTitle;
            }
            else if (instatnce.mainSubtitle.ContainsKey(name)) 
            {
                d = instatnce.mainSubtitle;
            }
            else if (instatnce.extraTitle.ContainsKey(name))
            {
                d = instatnce.extraTitle;
            }
            else if (instatnce.extraSubtitle.ContainsKey(name))
            {
                d = instatnce.extraSubtitle;
            }

            if (d != null)
            {
                if (d.ContainsValue(c))
                {
                    return true;
                }
            }

            return false;
        }
    }
}