using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace excelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\dosya yolunu yaz";

            Console.Write("Okumak istediğiniz sayfa indeksini girin (0'dan başlayarak): ");
            if (int.TryParse(Console.ReadLine(), out int selectedSheetIndex))
            {
                // Excel dosyasını FileStream ile açma
                using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    // FileStream'i IExcelDataReader ile sarma
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Tüm sayfaları alın
                        var dataSet = reader.AsDataSet();

                        // Seçilen sayfayı alın
                        if (selectedSheetIndex >= 0 && selectedSheetIndex < dataSet.Tables.Count)
                        {
                            var dataTable = dataSet.Tables[selectedSheetIndex];

                            // Verileri işleme
                            foreach (DataRow row in dataTable.Rows)
                            {
                                foreach (var item in row.ItemArray)
                                {
                                    Console.Write(item + "\t");
                                }
                                Console.WriteLine();
                            }
                        }
                        else
                        {
                            Console.WriteLine("Geçersiz sayfa indeksi.");
                        }
                    }
                }

                // Kullanıcıdan devam etmesini istemek için bir tuşa basmasını bekleyin
                Console.WriteLine("Devam etmek için bir tuşa basın...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Geçersiz sayfa indeksi girdiniz.");
                Console.ReadLine();
            }
        }
    }
}
