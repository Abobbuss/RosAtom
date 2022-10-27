using System;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace END
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application App = new Excel.Application();//ниже переменные для работы с Excel файлами
            Excel.Workbook Book1;
            Excel.Worksheet ListBook1;
            Excel.Workbook Book2;
            Excel.Worksheet ListBook2;

            Console.WriteLine("Здравствуйте!\nВведите полный путь к файлу otchet dlya sverki -2, пожалйуста");// нужно указывать с таким слешом - / !!
            string FirstNameFile = Console.ReadLine();
            Console.WriteLine("Введите полынй путь к файлу ostatki, пожалуйста");
            string SecondNameFile = Console.ReadLine();
            Console.WriteLine("Считывание файлов");

            Book1 = App.Workbooks.Open(FirstNameFile);
            ListBook1 = Book1.Worksheets["Sheet1"]; //открытие первого файла

            Book2 = App.Workbooks.Open(SecondNameFile); 
            ListBook2 = Book2.Worksheets["Sheet1"]; //открытие второго файла

            Range excelRange1 = ListBook1.UsedRange;
            Range excelRange2 = ListBook2.UsedRange;
            Console.WriteLine("Считывание файлов окончено\nНачинаем проверку!");

            for (int FirstNumberMaterial = 2; FirstNumberMaterial < excelRange1.Rows.Count; FirstNumberMaterial++)//цикл чтобы бегать по первому файлу
            {
                bool flag = true;// флаг для проверки и выхода из цикла
                if (ListBook1.Cells[FirstNumberMaterial, 7].Value != ListBook1.Cells[FirstNumberMaterial + 1, 7].Value && ListBook1.Cells[FirstNumberMaterial, 8].Value != 0)//исключение повторяющихся материалов, а так же исключение материалов с нулевым итогом
                {
                    int A = 2; //Начало файла
                    int B = excelRange2.Rows.Count - 1;//Конец файла
                    var X1 = ListBook1.Cells[FirstNumberMaterial, 7].Value;//Значение материала в данной точке
                    Int64 G1 = Convert.ToInt64(X1);//конвект.
                    for (int number = 0; ; number++) // цикл для автоматизация проверки ячеек между документами
                    {
                        var X = ListBook2.Cells[(B - A) / 2 + A, 4].Value;
                        Int64 G = Convert.ToInt64(X);
                        if (G == G1)
                        {
                            break;
                        }
                        if (G < G1)
                        {
                            A = (B - A) / 2 + A;
                        }
                        if (G > G1)
                        {
                            B = (B - A) / 2 + A;
                        }
                        if (number == 30)
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag == false)// условие для поиска строк, которых нет в файле остатки
                    {
                        Console.WriteLine("\nОшибка найдена!!!\nСтрочка 1 документа под номером " + FirstNumberMaterial + " ее не хватает в файле с остатками");
                        continue;
                    }
                    for (int SecondNumberMaterial = B = (B - A) / 2 + A; SecondNumberMaterial < excelRange2.Rows.Count; SecondNumberMaterial++)//цикл чтобы бегать по 2 файлу
                    {
                        if (ListBook1.Cells[FirstNumberMaterial, 7].Value == ListBook2.Cells[SecondNumberMaterial, 4].Value && ListBook2.Cells[SecondNumberMaterial, 4].Value != ListBook2.Cells[SecondNumberMaterial + 1, 4].Value)//исключение повторяющихся материалов во втором файле, а так же проверка, равен ли материал первого файла второму
                        {
                            if (ListBook1.Cells[FirstNumberMaterial, 8].Value == ListBook2.Cells[SecondNumberMaterial, 11].Value)
                            {
                                
                                break;
                            }
                            else 
                            {
                                Console.WriteLine("Найдено несовподение, строка первого документа под номером " + FirstNumberMaterial + " и второго документа под номером " + SecondNumberMaterial);
                                break;
                            }
                        }

                    }
                }

            }
            Console.WriteLine("\nИтоговые ошибки в том, что в файле с отстатками не хватает 2 строк из файла для проверки,если к итоговому значению файла для сверки прибавить найденные несостыковки, то получится итоговое значение в остатках, а значит все ошибки найдены!\nДля закрытия работы программы нажмите Enter.");
            Console.ReadLine();



        }
    }
}
