using System;

namespace Fox.Npoi.Test
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                DataShow.Basic();
                DataShow.Basic2();
                DataShow.Basic3();
                DataShow.Basic4();
                DataShow.Basic5();
                DataShow.Basic6();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.Read();
            }
        }
    }
}