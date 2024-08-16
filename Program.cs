using System.Text;

namespace FileWork
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "Test.txt";


            if (File.Exists(path)) 
            {
                int lines = File.ReadAllLines(path).Length;    

                File.AppendAllText(path, $"\n{++lines}: Neue Zeile?");
                File.CreateSymbolicLink("TestLink.txt", path);   

            } 
            else 
            {
                FileStream fs = File.Create("test.txt");
                fs.Write(Encoding.UTF8.GetBytes("Hello World!"));
                fs.Close();
            }

            

        }
    }
}
