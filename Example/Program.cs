using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GenericExporter;
namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            var exporter = new Exporter();
            var emptyList = new List<User>();
            var items = new List<User>();
            for (var i = 0; i < 100; i++)
                items.Add(new User
                {
                    Id = i,
                    FirstName = $"FirstName{i}",
                    LastName = $"LastName{i}",
                    CreatedDate = GetRandomDate(),
                    Salary = new Random().NextDouble()
                });
            var result = exporter.Export(items.Select(x => x));
            File.WriteAllBytes("Example.xlsx", result);
            var result1 = exporter.Export(GetDynamics(items));
            File.WriteAllBytes("Example1.xlsx", result1);
            var result2 = exporter.Export(GetDynamics(emptyList));
            if(result2==null)
            {
                File.Delete("ExampleEmpty.xlsx");
            }
            else
                File.WriteAllBytes("ExampleEmpty.xlsx", result2);

        }
        private static DateTime GetRandomDate()
        {
            Random gen = new Random();
            DateTime start = new DateTime(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range));
        }
        private static IEnumerable<dynamic> GetDynamics(IEnumerable<User> items)
        {
            return items.Select(x => new
            {
                x.Id,
                x.FirstName,
                x.LastName,
                x.CreatedDate,
                x.Salary
            });
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
