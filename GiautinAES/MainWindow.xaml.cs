using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using System.IO;
using System.Drawing;
using Microsoft.Win32;
using System.Security.Cryptography;
using System.Threading.Tasks.Dataflow;
using System.Collections.Specialized;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;

namespace GiautinAES
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // đường dẫn file ảnh
        private string inPath = "";
        private string cipherText; // bản mã sau khi mã hóa
        private string hiddenMessage; // thông điệp được giấu trong ảnh
        private string key = ""; // khóa dùng cho mã hóa và giải mã
        private byte[] IV = {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16};

        private static Random random = new Random();

        public MainWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.txtNews.IsEnabled = false;
        }

        private void btn_selectImg(object sender, RoutedEventArgs e)
        {
            // chọn ảnh để giấu tin

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                inPath = openFileDialog.FileName;
                string extension = System.IO.Path.GetExtension(inPath);
                if (extension == ".png" || extension == ".bmp")
                {
                    this.rootImg.Source = new BitmapImage(new Uri(new Uri(Directory.GetCurrentDirectory(), UriKind.Absolute), new Uri(inPath, UriKind.Relative)));
                }
                else
                {
                    MessageBox.Show("File ảnh phải là .png hoặc .bmp");
                }
            }
        }

        private void btn_start(object sender, RoutedEventArgs e)
        {

            // Xử lý textbox rỗng
            if (this.txtContent.Text == "")
            {
                MessageBox.Show("Nhập nội dung cần mã hóa");
                this.txtContent.Focus();
            }else if (this.txtPassword.Text == "")
            {
                MessageBox.Show("Bạn phải nhập mật khẩu trước");
                this.txtPassword.Focus();
            }else if(this.txtPassword.Text.Length != 16)
            {
                MessageBox.Show("Mật khẩu phải đủ 16 ký tự");
                this.txtPassword.Text = "";
                this.txtPassword.Focus();
            }else
            {
                try
                {
                    using (Aes aes = Aes.Create())
                    {
                        string plainText = this.txtContent.Text;

                        // chuyển khóa từ byte sang string dùng giải mã
                        aes.KeySize = 128;
                        //key = this.txtPassword.Text;
                        aes.Key = Encoding.UTF8.GetBytes(this.txtPassword.Text);
                        // tạo IV
                        aes.IV = IV;
                        //aes.GenerateIV();
                        // chuyển IV từ byte sang string
                        //IV = Convert.ToBase64String(aes.IV);
                        // tạo đối tượng mã hóa
                        ICryptoTransform encryptor = aes.CreateEncryptor();
                        byte[] encryptedData;
                        using (MemoryStream msEncrypt = new MemoryStream())
                        {
                            using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                            {
                                using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                                {
                                    swEncrypt.Write(plainText);
                                }
                                encryptedData = msEncrypt.ToArray();
                            }
                        }
                        cipherText = Convert.ToBase64String(encryptedData);
                    }
                    MessageBox.Show("Mã hóa thành công");
                } catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi khi mã hóa: " + ex.Message);
                }
            }
        }

        private void btn_download(object sender, RoutedEventArgs e)
        {
            // bắt đầu giấu tin

            Bitmap bitmapImg = new Bitmap(inPath);
            for (int i = 0; i < bitmapImg.Width; i++)
            {
                for (int j = 0; j < bitmapImg.Height; j++)
                {
                    System.Drawing.Color pixel = bitmapImg.GetPixel(i, j);
                    if (i < 1 && j < cipherText.Length)
                    {
                        Console.WriteLine("R[" + i + "][" + j + "] : " + pixel.R);
                        Console.WriteLine("R[" + i + "][" + j + "] : " + pixel.G);
                        Console.WriteLine("R[" + i + "][" + j + "] : " + pixel.B);

                        char letter = Convert.ToChar(cipherText.Substring(j, 1));
                        int value = Convert.ToInt32(letter);
                        Console.WriteLine("Letter: " + letter + "\n Value: " + value);
                        bitmapImg.SetPixel(i, j, System.Drawing.Color.FromArgb(pixel.R, pixel.G, value));
                    }
                    if (i == bitmapImg.Width - 1 && j == bitmapImg.Height - 1)
                    {
                        bitmapImg.SetPixel(i, j, System.Drawing.Color.FromArgb(pixel.R, pixel.G, cipherText.Length));
                    }
                }
            }

            // lưu ảnh đã giấu tin
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PNG|*.png|BMP|*.bmp";
            saveFileDialog.ShowDialog();
            string fileName = saveFileDialog.FileName;
            bitmapImg.Save(fileName);

        }

        private void btn_slDecryptImg(object sender, RoutedEventArgs e)
        {
            // chọn ảnh đã giấu tin từ trước

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                inPath = openFileDialog.FileName;
                string[] fileTail = inPath.Split('.');
                if (fileTail[1] == "png" || fileTail[1] == "bmp")
                {
                    this.decryptImg.Source = new BitmapImage(new Uri(new Uri(Directory.GetCurrentDirectory(), UriKind.Absolute), new Uri(inPath, UriKind.Relative)));
                }
                else
                {
                    MessageBox.Show("File ảnh phải là .png hoặc .bmp");
                }
            }
        }

        private void btn_getMessage(object sender, RoutedEventArgs e)
        {
            if (this.txtPw.Text == "")
            {
                MessageBox.Show("Bạn phải nhập mật khẩu trước");
                this.txtPw.Focus();
            }else
            {
                // lấy thông điệp từ ảnh
                try
                {
                    Bitmap bitmapImg = new Bitmap(inPath);
                    System.Drawing.Color pixel = bitmapImg.GetPixel(bitmapImg.Width - 1, bitmapImg.Height - 1);
                    int messLength = pixel.B;
                    for (int i = 0; i < bitmapImg.Width; i++)
                    {
                        for (int j = 0; j < bitmapImg.Height; j++)
                        {
                            System.Drawing.Color pixel1 = bitmapImg.GetPixel(i, j);
                            if (i < 1 && j < messLength)
                            {
                                Console.WriteLine("-------------");
                                Console.WriteLine("R[" + i + "][" + j + "] : " + pixel1.R);
                                Console.WriteLine("R[" + i + "][" + j + "] : " + pixel1.G);
                                Console.WriteLine("R[" + i + "][" + j + "] : " + pixel1.B);

                                int value = pixel1.B;
                                Console.WriteLine("Value: " + value);
                                char c = Convert.ToChar(value);

                                string letter = c.ToString();
                                hiddenMessage += letter;
                            }
                        }
                    }
                    MessageBox.Show("Lấy thông điệp từ ảnh thành công");
                } catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi: " + ex.Message);
                }
            }
        }

        private void btn_decrypt(object sender, RoutedEventArgs e)
        {
            // giải mã thông điệp lấy được từ ảnh

            // chuyển thông điệp sang mảng byte
            try
            {
                byte[] bytes = Convert.FromBase64String(hiddenMessage);
                string plainText = "";

                using (Aes aes = Aes.Create())
                {
                    aes.Key = Encoding.UTF8.GetBytes(this.txtPw.Text);
                    //aes.IV = Convert.FromBase64String(IV);
                    aes.IV = IV;
                    // tạo đối tượng giải mã
                    ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                    // Create the streams used for decryption.
                    using (MemoryStream msDecrypt = new MemoryStream(bytes))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                            {

                                // Read the decrypted bytes from the decrypting stream
                                // and place them in a string.
                                plainText = srDecrypt.ReadToEnd();
                            }
                        }
                    }
                }
                txtNews.Text = plainText;
            } catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi giải mã: " + ex.Message);
            }
        }

        private void btn_selectFileInput(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "TXT|*.txt";
            if(openFile.ShowDialog() == true)
            {
                string filePath = openFile.FileName;
                string extension = System.IO.Path.GetExtension(filePath);
                if (extension == ".txt")
                {
                    this.txtContent.Text = File.ReadAllText(openFile.FileName);
                }else
                {
                    MessageBox.Show("Chức năng đang được phát triển");



                    //this.txtContent.Text = readDocX(filePath);
                }
            }
        }

        private string readDocX(string fileName)
        {
            Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.Application();
            
            object file = fileName;
            object readOnly = true;
            object addToRecentFiles = false;

            Microsoft.Office.Interop.Word._Document document = app.Documents.Open(ref file, ReadOnly: readOnly, AddToRecentFiles: addToRecentFiles);
            string data = document.Content.Text;

            object saveChanges = false;
            app.Quit(SaveChanges: saveChanges);

            return data;
        }

        private static string randomString(int length)
        {
            const string chars = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void btn_genKey(object sender, RoutedEventArgs e)
        {
            Random random = new Random();
            this.txtPassword.Text = randomString(16);
        }

        private void countChar(object sender, TextChangedEventArgs e)
        {
            int count = this.txtPassword.Text.Length;
            this.count.Content = (count + "/16").ToString();
        }

        private void countChar1(object sender, TextChangedEventArgs e)
        {
            int count = this.txtPw.Text.Length;
            this.count1.Content = (count + "/16").ToString();
        }
    }
}
