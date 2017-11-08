using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;

namespace Quack
{
    public interface ISomeMore
    {
        void MoreStuff();
    }

    public interface IBaseQuack : IDisposable
    {
        void RootStuff();
    }

    public interface IQuack : IBaseQuack, ISomeMore
    {
        string Name { get; }

        bool Visible { get; set; }

        void Talk();

        string Shit(int x, ref bool any);
    }

    public class Quack
    {
        public string Name
        {
            get
            {
                return "Quack11";
            }
        }

        public bool Visible { get; set; }

        public string Shit(int x, ref bool any)
        {
            any = true;
            return "x " + x.ToString();
        }

        public void Talk()
        {
            Console.WriteLine("Quack Quack Quack");
        }

        public void RootStuff()
        {
            Console.WriteLine("RootStuff");
        }

        public void MoreStuff()
        {

        }

        public void Dispose()
        {

        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DuckTypingProxyFactory factory = new DuckTypingProxyFactory();
            Quack cuacky = new Quack();
            IQuack proxy = factory.GenerateProxy<IQuack>(cuacky);
            Speak(proxy);
            Console.ReadLine();
        }

        static void Speak(IQuack bird)
        {
            Console.WriteLine(bird.Name);
            bird.Talk();
        }
    }
}
