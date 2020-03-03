using System;
using System.Collections.Generic;

namespace NExportToExcel
{
    [ExportToExcelAttribute("اسامی زاران اعزامی")]
    public class Person
    {
        [ExportToExcelAttribute("آی دی")]
        public int Id { get; set; }
        public int Order { get; set; }
        [ExportToExcelAttribute("نام")]
        public string FirstName { get; set; }
        [ExportToExcelAttribute("نام خانوادگی")]
        public string LastName { get; set; }
        //[ShouldBeMerged(nameof(Order))]
        [ExportToExcelAttribute("آیتمی که باید مرچ شود")]
        public int MergePropery { get; set; }
    }
    class Program
    {

        static void Main(string[] args)
        {
            var persons = new List<Person>();
            for (int i = 0; i < 10; i++)
            {
                persons.Add(new Person
                {
                    FirstName = "نام " + i,
                    LastName = "نام خانوادگی " + i,
                    Id = i,
                    //Order = i < 3 ? 0
                    //: i < 5 ? 1
                    //: 2,
                    Order = i < 2 ? 0 : i < 5 ? 1 : 2,
                    MergePropery = i
                });
            }

            persons.ExportToExcel(@"D:\sources\NExportToExcel");
            Console.ReadKey();
        }
    }
}
