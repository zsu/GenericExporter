using System;
using System.Collections.Generic;
using System.IO;
using GenericExporter;
namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            var exporter = new Exporter();
            var items = new List<User>();
            for(var i=0;i<100;i++)
                items.Add(new User
                {
                    Id = i,
                    FirstName = $"FirstName{i}",
                    LastName = $"LastName{i}",
                    CreatedDate = GetRandomDate(),
                    Salary = new Random().NextDouble()
                });
            var result=exporter.Export(items);
            File.WriteAllBytes("Example.xlsx", result);

        }
        private static DateTime GetRandomDate()
        {
            Random gen = new Random();
            DateTime start = new DateTime(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range));
        }

    }
     class User
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime CreatedDate { get; set; }
        public double Salary { get; set; }
    }
}
