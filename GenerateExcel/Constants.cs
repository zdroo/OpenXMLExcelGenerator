using System.Collections.Generic;
using System.Reflection;
using System;

public static class Constants
{
    public enum ChartType
    {
        ClusteredColumnBarChart,
        LineChart,
        ScatteredChart
    }
    public static List<string> Alphabet = new List<string>()
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
        };

    public static string ApplicationPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    public static int ExcelNumber = 1;
    public static int PictureNumber = 1;

    public const string LightGrayBgCode = "F6F8FF";
    public const string ChartTitleBlueCode = "021146";
    public const string SigtestGreenCode = "58FF8A";
    public const string SigtestYellowCode = "FFEA30";
    public const string BlueGrayCode = "ADC1CB";
    public const string AxaBlueCode = "00008F";
    public const string HeaderDarkBlueCode = "000047";
    public const string ChapterNumAquaCode = "00AEC6";
    public const string LightGrayCode = "9E9ED";
    public const string DarkBlueGrayCode = "4F5F75";
    public const string DarkGrayCode2 = "4F5D69";
    public const string NavyBlueCode = "243B5F";
    public const string OutlineGrayCode = "D9D9D9";
    public const string LightPinkCode = "E196AA";
    public const string DarkTurquoisCode = "027180";
    public const string DarkRedCode = "914146";
    public const string LightBlueCode = "00AEC6";
    public const string ScatteredLightGrayCode = "BFBFBF";
    public const string BgGrayCode = "F0F0F0";
    public const string MustardYellowCode = "C2A10D";
    public const string MarineBleuCode = "7FA2B1";
    public const string CyanCode = "1DE1FC";
    public const string TahitiBlueCode = "00AEC6";
    public const string UpliftsGreenCode = "10BA38";
    public const string SlideBgGrayCode = "F1F2F6";
    public const string WhiteCode = "FFFFFF";
    public const string BlackCode = "000000";

    public const Int64 Pixel = 12600;

    #region Binary Data
    public static string RedRankingArrowData = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAkCAYAAACe0YppAAAABHNCSVQICAgIfAhkiAAAAh9JREFUWEftlr9rE2EYx5/n0twpgnSqi9VFcJG2l1AolA7SQRBbzCXVzUqF9i/oItihHQoO7jpZp2J7aUEodBBEHErEXLI5OKij2OJSenfh3qfPRUzT5O7eS9JOXpYE7vt+P5fnx/cOIcbH1o1PiDgeQwpqWelD2PBkWpQJ/OsJOKpKSanDqpMMV2BlknVK1qm5AkmAJAHyrwL/YWS6uvGEEAajQoEA5hAwUtM4T7AMSHwkdKxq6i/lBdpD0zexL11i2eU4r0E9azx4qFXNt/Ue1zKFCQH0nn+mezaOMhD0VKsUV31JY7ic4fwDUGgd+HXyPOAE9OpCubgQONV2Jr/I1OdnDibYUS3zHns3et/27+yM8ZIHaf7s4LSnHhzcxu8f7GbPNjDfEroZw+SvXK9wnu1vGnmjWNn+0+oV2E/KZtMuXf/I8LFu4dzT3xqmsvhl42eQR+gg0cj9fgdTn3nUbnQMJzhEARNq1bS6ekhQduaaI7wSD/qVDuCeQngnbW366xkeIzJDZyh3C1LKHi/eJZm2vp8kZlVr641MG2tna3phUiDtslkq0pBoSbOKKzLoqQCRiV0994hQWQvVEaxplvlY5hMYILJDjm4842RbbtP9DYgpLp+QeXQF9g85ev4193v2BEBlDojx1oCQ3UCsHjebcMAorp5/x/C79YA4EmP4dWtfBmq93jHYN6CrMxedAbGtEc2jVfzRKdTXHwOSTvOe4oZoSwAAAABJRU5ErkJggg==";
    public static string GreenRankingArrowData = "iVBORw0KGgoAAAANSUhEUgAAAB0AAAAkCAYAAAB15jFqAAAABHNCSVQICAgIfAhkiAAAAbdJREFUWEftlbFKA0EQhmcSU8TUQoq8gE3AQrASQVIIYiGEmMJGbGxEJMm1AUG4RBAbW8skxEawEKxsRFCwUx/ASqwEEZRk/DckwUsut7tXyl15NzPfzu58e0whnoLUVpG21OJKKUQ6sW1SUerZLskdEpNCstVi58y2hhV0XdxMgvgeSWkFEqJOjGSlwc61DdgYuin11Dd1AeTZEcAnOp5Hx8+mYCNoVaqxF0peEXHOv7C8MsXnGlx6NwEbQTE4pwjcCS4oDx1KLLZ5/0sH1kIL4u5hS491hfpnfNmi8hox47gnP4FQpQaTXGBbYybQfsxRkyvlUNC/algAe6E6lXw7zUstjdYeB2rYQ5VKtNzgyo1f7hhUqfFDcovgrC1sJP4DHS/4qeSB6tWwXYa/Sh6omRrW4DGVhtANcXcxpSe2JU3i4Y9HpR60KG4OH9SNY6OGCW8Yg/o1/JUc9YKVGjhwNTgpqyohggcqMc5xGxdAZlINrHAaN1Kg7INcxLZR6ylgPZ0mOwfaazAvhzNxmnoza0zyKHqui42gwx2Kttd/WKJB8u5LpEykjOZajZSJlJkwIv/mL/MLusX6Jdg6lTIAAAAASUVORK5CYII=";

    #endregion

}