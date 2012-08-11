/*********************************************************************************
 *Copyright 2012 Michael Braun                                                   *
 *Licensed under the Red Spider Project License.                                 *
 *See the License.txt that shipped with your copy of this software for details.  *
 *Version 1.0.0                                                                  *
 *********************************************************************************/
using System;
using System.Net;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace xkcdDesktop
{

    class Program
    {
        /// <summary>
        /// Start the program.
        /// </summary>
        /// <param name="arrgh">arrgh like a pirate</param>
        static void Main(string[] args)
        {
            Console.WriteLine(@"Type ? or help for help. Copyright 2012 Michael Braun");
            Program pg = new Program();
            pg.CheckForCommands();
        }

        private void CheckForCommands()
        {
            while (true)
            {
                int comicNumber;
                string command = Console.ReadLine();
                if (int.TryParse(command, out comicNumber))
                {
                    DisplayComic(comicNumber);
                }
                else
                {
                    switch (command.ToLower())
                    {
                        default:

                            continue;
                        case "?":
                        case "help":
                            Console.WriteLine(@"List of commands: (All commands should end by pressing the enter key.)
   r         -Gets a random xkcd comic and displays it to the desktop
   #         -Gets the xkcd comic of the number which you specify and displays 
              it on the desktop. (ex. typing 1024 would get comic number 1024.)
   ? or help -Displays help for commands
   clear     -clears all text from console and changes to default colors.
   q         -exits the application");
                            continue;
                        case "clear":
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Console.BackgroundColor = ConsoleColor.Black;
                            Console.Clear();
                            continue;
                        case "q":
                            break;
                        case "r":
                            DisplayComic();
                            continue;
                    }
                }
            }
        }

        /// <summary>
        /// Where everything comes together. Displays comic and information.
        /// </summary>
        /// <param name="rand">Random to get a comic.</param>
        /// <param name="wc">WebClient to get comic.</param>
        private void DisplayComic()
        {
            WebClient wc = new WebClient();
            Random rand = new Random();
            int desktopStrip1 = rand.Next(0, FindCurrentStripNumber(wc));
            //Set strip.
            //desktopStrip1 = 126;
            string JSON = wc.DownloadString(new Uri("http://xkcd.com/" + desktopStrip1.ToString() + "/info.0.json"));
            switch (desktopStrip1)
            {
                default:
                    Console.WriteLine("Number: " + desktopStrip1.ToString());
                    Console.WriteLine("Title: " + FindTitle(JSON) + "\n");
                    Console.WriteLine("Alt text: " + FindAltText(JSON) + "\n");
                    SaveToDirectory(wc, JSON);
                    break;
                case 404:
                    FourOFour();
                    break;
                //These next are all red spider themed.
                case 8:
                case 43:
                case 47:
                case 126:
                case 427:
                    RedSpiderComic();
                    goto default;
            }
            wc.Dispose();
        }

        /// <summary>
        /// Displays a strip on the desktop.
        /// </summary>
        /// <param name="desktopStrip">Which strip to display.</param>
        private void DisplayComic(int desktopStrip)
        {
            WebClient wc = new WebClient();
            string JSON = wc.DownloadString(new Uri("http://xkcd.com/" + desktopStrip.ToString() + "/info.0.json"));
            switch (desktopStrip)
            {
                default:
                    Console.WriteLine("Number: " + desktopStrip.ToString());
                    Console.WriteLine("Title: " + FindTitle(JSON) + "\n");
                    Console.WriteLine("Alt text: " + FindAltText(JSON) + "\n");
                    SaveToDirectory(wc, JSON);
                    break;
                case 404:
                    FourOFour();
                    break;
                //These next are all red spider themed.
                case 8:
                case 43:
                case 47:
                case 126:
                case 427:
                    RedSpiderComic();
                    goto default;
            }
            wc.Dispose();
        }

        /// <summary>
        /// Converts the comic and saves it as a bitmap.
        /// </summary>
        /// <param name="filename">Where you want to put the file.</param>
        /// <param name="comicFilename">Filename of unconverted comic.</param>
        private void SaveAsBitmap(string filename, string comicFilename)
        {
            Image myImage = Image.FromFile(comicFilename);
            using (Bitmap b = new Bitmap(myImage.Width, myImage.Height))
            {
                using (Graphics g = Graphics.FromImage(b))
                {
                    g.Clear(Color.White);
                    g.DrawImageUnscaled(myImage, 0, 0);
                }
                myImage.Save(filename + "comic.bmp", ImageFormat.Bmp);
            }
        }

        /// <summary>
        /// Gets image and downloads it to the current directory of the program (debug).
        /// </summary>
        /// <param name="wc">Webclient to download the JSON code.</param>
        /// <param name="JSON">The JSON code to parse.</param>
        private void SaveToDirectory(WebClient wc, string JSON)
        {
            String comicDirectory = Directory.GetCurrentDirectory();
            String comicFromWebDir = Directory.GetCurrentDirectory() + "\\comic.jpg";
            File.SetAttributes(comicDirectory, FileAttributes.Normal);//to prevent Unauthorized Access Exception
            Uri imgURI = new Uri(FindImageURI(JSON).Trim((new char[2] { '\\', '"' })));
            try
            {
                wc.DownloadFile(imgURI, comicFromWebDir);
            }
            catch (WebException e)//I guess people will have problems with connectivity
            {
                Console.WriteLine("Web Error:" + e.Message);
            }
            SaveAsBitmap(comicDirectory, comicFromWebDir);
            Wallpaper.SetDesktopWallpaper(WallPaperType.Fit, comicDirectory + "comic.bmp");
        }

        /// <summary>
        /// Finds the current comic strip number by using xkcd.com
        /// </summary>
        /// <param name="wc">Webclient to retrieve JSON.</param>
        /// <returns>Number of the current comic.</returns>
        private int FindCurrentStripNumber(WebClient wc)
        {
            string currentStripJSON = wc.DownloadString(new Uri("http://xkcd.com/info.0.json"));
            int currStripNumIndex = currentStripJSON.IndexOf("\"num\": ");
            int currentStrip = int.Parse(currentStripJSON.Substring(currStripNumIndex + 7, 4));
            return currentStrip;
        }

        /// <summary>
        /// Finds the comic's alt text.
        /// </summary>
        /// <param name="JSON">JSON code from the comic.</param>
        /// <returns>Alt text of comic.</returns>
        private string FindAltText(string JSON)
        {
            int altTextIndex;
            int altTextLength;
            string altText;
            if (JSON.Contains("\"Alt text\":"))
            {
                altTextIndex = JSON.IndexOf("\"Alt text\": ");
                altTextLength = (-altTextIndex - 12 + JSON.IndexOf('}', altTextIndex));
                altText = JSON.Substring(12 + altTextIndex, altTextLength);
            }
            else if (JSON.Contains("\"alt\":"))
            {
                altTextIndex = JSON.IndexOf("\"alt\": ");
                altTextLength = (-altTextIndex - 7 + JSON.IndexOf("\",", altTextIndex));
                altText = JSON.Substring(7 + altTextIndex, altTextLength);
            }
            else
            {
                altTextIndex = -10;
                altText = string.Empty;
            }
            return altText;
        }

        /// <summary>
        /// Finds the comic's title.
        /// </summary>
        /// <param name="JSON">JSON code to extract the title from.</param>
        /// <returns>Comic's title.</returns>
        private string FindTitle(string JSON)
        {
            int titleIndex;
            int titleLength;
            string titleText;
            if (JSON.Contains("\"safe_title\":"))
            {
                titleIndex = JSON.IndexOf("\"safe_title\":");
                titleLength = (-titleIndex - 13 + JSON.IndexOf(',', titleIndex));
                titleText = JSON.Substring(13 + titleIndex, titleLength);
                return titleText;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Finds the directory of the image so you can retrieve it.
        /// </summary>
        /// <param name="JSON">JSON code for comic.</param>
        /// <returns>The location of the image.</returns>
        private string FindImageURI(string JSON)
        {
            int imgIndex = JSON.IndexOf("\"img\": ");
            int imgLength = -7 - imgIndex + JSON.IndexOf(',', imgIndex);
            string imgURIString = JSON.Substring(imgIndex + 7, imgLength).Replace("\\", string.Empty);
            return imgURIString;
        }

        /// <summary>
        /// In case of comic 404.
        /// </summary>
        private void FourOFour()
        {
            Console.CursorVisible = false;
            Thread.Sleep(1000);
            for (int i = 0; i < 3; i++)
            {
                Console.Write("5");
                Thread.Sleep(100);
            }
            Console.Write('-');
            for (int i = 0; i < 4; i++)
            {
                Console.Write("5");
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            Console.Clear();
            Thread.Sleep(2000);
            Console.WriteLine("Accessing Server:");
            for (int i = 0; i < 53; ++i)
            {
                Console.Write("\r{0}%   ", i);
                Thread.Sleep(20);
            }
            Thread.Sleep(100);
            for (int i = 53; i < 80; i++)
            {
                Console.Write("\r{0}%   ", i);
                Thread.Sleep(10);
            }
            Thread.Sleep(1000);
            Console.Write("\r78%   ");
            for (int i = 78; i < 100; i++)
            {
                Console.Write("\r{0}%   ", i);
                Thread.Sleep(100);
            }
            for (int i = 0; i < 9; i++)
            {
                Console.Write('.');
                Thread.Sleep(500);
            }
            Console.Write("\r 100%");
            Thread.Sleep(500);
            Console.Clear();

            Console.WriteLine("Pinging 127.0.0.1 with 16-bit WAN connection:\n");
            Thread.Sleep(52);
            Console.WriteLine(@"Reply from 127.0.0.1: data-stream=open time=52ms <WOC protocol=enabled>");
            Thread.Sleep(48);
            Console.WriteLine(@"Reply from 127.0.0.1: data-stream=open time=48ms <WOC protocol=disabled>");
            Thread.Sleep(92);
            Console.WriteLine(@"Reply from 127.0.0.1: data-stream=open time=92ms <WOC protocol=enabled>\n\n");
            Thread.Sleep(400);
            Console.WriteLine(@"Ping Statistics for 127.0.0.1
    Packets sent: 4 Recieved: 3 Lost: 1
Approx. round trip-times in parsecs
    Input: 43 Output: 427 Average: 230" + "\n\n");
            for (int i = 0x0; i < 0xFFF; i++)
            {
                Console.WriteLine("0x10000{0:X}", i);
            }
            Thread.Sleep(500);
            Console.Clear();
            Thread.Sleep(1000);
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.Clear();
            ConsoleHelper.SetConsoleIcon(SystemIcons.Error);
            ConsoleHelper.SetConsoleFont(5);
            Console.WriteLine("404 ERROR");
        }

        /// <summary>
        /// When a comic features red spiders, this method is called.
        /// </summary>
        private void RedSpiderComic()
        {
            Console.BackgroundColor = ConsoleColor.Gray;
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Red;
            string asciiArt = @"                                                                                
                      $$$$     $$$#$$$$$                                        
                   $$$$$$     $$$$$     $$                                      
                 ®$$$$®$$®$$$$$$$$$$®$   $$                                     
                ®$$$$$#$$$#$$$$$$$$$$$$$  $$$$®    ,                            
               ®®$$$$$$®®$$$$$$$$$$$$$$$$  ,$$                                  
               ®®®$$$$$$$$®#$$#$$$$$$$$$$$,    $$$$®o $$                        
              ,®®®®®$$$$$$$$$$®®$$$$$$$$$$$      $$     ,                       
              ,,®®®®®®®®®$$$$$$$$®$$$$$$$$$$$     $$$$ ,,$,$$                   
        $$$$$$$$$$$$$®®®®®®®®$$$$$®$$$$$$$$$$$,  $$$$$$$     $                  
        $      ,,,®®®®®®®®®$$$$$$$$$$$$®$$$$$$$$$$$$$$$$$    $ o                
       ,$        ,,,o$,,$$$®®®$$$$$$$$$$$®$$$$$$$$$$$$$$$$,,$,,$$$$ ,           
       ,      $$$®o ,,,,,®®®®®®$$$$$$$$$$$$$$$$$$$$$$$$$$$$       $  $,         
       $      $        ,,,,®®® $®®$$$$$$$$$$$$$$$$$$$$$$$$$$           $        
       $      $   ,,,,, $,$$$®®®®®®$$$$$$$$$$$$$$$$$$$##$$$$             $      
           ,,,,,    $$$    ,,,,®®®®®®$$$$$$®# #$$$$$$$$$$$$                ,    
      ,,,    ,      $       ,,,,,®®,$$®®$$$$###$$$$$$$$$$$,                     
             $      $   ,,,    ,$$®®®®®®®®®$$$$$$$$$$$$$$                       
             $   ,,,,        ,$ ,,,,,,$®®®®®®®®®®®®®®®                          
             ,,             ,,,      ,,,,,,,,,,,,                               
                   $      ,,         ,,                                         
                   $  ,,   ,       ,                                            
                   $,      o      ,                                             
                           $    ,                                               
                           $  ,                                                 
            ";
            char[] asciiArray = asciiArt.ToCharArray();
            foreach (char c in asciiArray)
            {
                switch (c)
                {
                    case '®':
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.Write('*');
                        break;
                    case '#':
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.Write('#');
                        break;
                    case '$':
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write('$');
                        break;
                    case ',':
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.Write(',');
                        break;
                    case 'o':
                        Console.ForegroundColor = ConsoleColor.DarkGray;
                        Console.Write('o');
                        break;
                    case '\n':
                        Console.Write('\n');
                        break;
                    case ' ':
                        Console.Write(' ');
                        break;
                }
            }
        }
    }


    /// <summary>
    /// Good for changing fonts and icons.
    /// </summary>
    public static class ConsoleHelper
    {
        [DllImport("kernel32")]
        public static extern bool SetConsoleIcon(IntPtr hIcon);

        public static bool SetConsoleIcon(Icon icon)
        {
            return SetConsoleIcon(icon.Handle);
        }

        [DllImport("kernel32")]
        private extern static bool SetConsoleFont(IntPtr hOutput, uint index);

        private enum StdHandle
        {
            OutputHandle = -11
        }

        [DllImport("kernel32")]
        private static extern IntPtr GetStdHandle(StdHandle index);

        public static bool SetConsoleFont(uint index)
        {
            return SetConsoleFont(GetStdHandle(StdHandle.OutputHandle), index);
        }

        [DllImport("kernel32")]
        private static extern bool GetConsoleFontInfo(IntPtr hOutput, [MarshalAs(UnmanagedType.Bool)]bool bMaximize,
            uint count, [MarshalAs(UnmanagedType.LPArray), Out] ConsoleFont[] fonts);

        [DllImport("kernel32")]
        private static extern uint GetNumberOfConsoleFonts();

        public static uint ConsoleFontsCount
        {
            get
            {
                return GetNumberOfConsoleFonts();
            }
        }

        public static ConsoleFont[] ConsoleFonts
        {
            get
            {
                ConsoleFont[] fonts = new ConsoleFont[GetNumberOfConsoleFonts()];
                if (fonts.Length > 0)
                    GetConsoleFontInfo(GetStdHandle(StdHandle.OutputHandle), false, (uint)fonts.Length, fonts);
                return fonts;
            }
        }

    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct ConsoleFont
    {
        public uint Index;
        public short SizeX, SizeY;
    }

    /// <summary>
    /// To set the wallpaper, you use this class to deal with the registry/
    /// </summary>
    public static class Wallpaper
    {
        public static bool SupportFitFillWallpaperStyles
        {
            get
            {
                return (Environment.OSVersion.Version >= new Version(6, 1));
            }
        }

        public static void SetDesktopWallpaper(WallPaperType style, string path)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
            switch (style)
            {
                case WallPaperType.Tiled:
                    key.SetValue(@"WallpaperStyle", "0");
                    key.SetValue(@"TileWallpaper", "1");
                    break;
                case WallPaperType.Center:
                    key.SetValue(@"WallpaperStyle", "0");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case WallPaperType.Stretch:
                    key.SetValue(@"WallpaperStyle", "2");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case WallPaperType.Fit: //Windows 7 and later
                    if (!SupportFitFillWallpaperStyles)
                        goto case WallPaperType.Stretch;
                    key.SetValue(@"WallpaperStyle", "6");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case WallPaperType.Fill: //Windows 7 and later
                    if (!SupportFitFillWallpaperStyles)
                        goto case WallPaperType.Stretch;
                    key.SetValue(@"WallpaperStyle", "10");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
            }
            key.Close();

            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        }

        private const uint SPI_SETDESKWALLPAPER = 20;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDWININICHANGE = 0x02;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, string pvParam, uint fWinIni);
    }

    /// <summary>
    /// How do you want the wallpaper to be arranged?
    /// </summary>
    public enum WallPaperType
    {
        Tiled,
        Center,
        Stretch,
        Fit,
        Fill,
    }
}
