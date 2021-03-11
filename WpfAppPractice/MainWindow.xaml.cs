using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Drawing;

namespace WpfAppPractice
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            userImage.Source = new BitmapImage(new Uri(@"Z:\WpfAppPractice\WpfAppPractice\Resources\babyYoda.jpeg"));
        }

        private void ImageToBytesBtn_Click(object sender, RoutedEventArgs e)
        {
            //userImageBytes.Text = ImageToByteArray(System.Drawing.Image.FromFile(@"Z:\WpfAppPractice\WpfAppPractice\Resources\babyYoda.jpeg"))[0].ToString();
            userImageBytes.Text = ImageToByteArr(userImage.Source as BitmapImage)[0].ToString();

            //userImageFromBytes.Source = ByteArrayToBitmapImage(ImageToByteArray(System.Drawing.Image.FromFile(@"Z:\WpfAppPractice\WpfAppPractice\Resources\babyYoda.jpeg")));
            userImageFromBytes.Source = ByteArrayToBitmapImage(ImageToByteArr(userImage.Source as BitmapImage));

        }

        public byte[] ImageToByteArray(System.Drawing.Image image)
        {
            return (byte[])(new ImageConverter()).ConvertTo(image, typeof(byte[]));
        }

        public byte[] ImageToByteArr(BitmapImage image)
        {
            JpegBitmapEncoder encoder = new JpegBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(image));

            using (MemoryStream ms = new MemoryStream())
            {
                encoder.Save(ms);

                return ms.ToArray();
            }
        }

        public BitmapImage ByteArrayToBitmapImage(byte[] imageBytesArray)
        {
            using (var ms = new MemoryStream(imageBytesArray))
            {
                var bitmapImg = new BitmapImage();
                bitmapImg.BeginInit();
                bitmapImg.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImg.StreamSource = ms;
                bitmapImg.EndInit();

                return bitmapImg;
            }
        }
    }
}
