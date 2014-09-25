using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Random
{
    class Random
    {
        int studentnum;
        private ArrayList result = new ArrayList();
        private ArrayList haven = new ArrayList();
        private static Random instance;
        System.Random random = new System.Random((DateTime.UtcNow.Millisecond*29+7) % 1000);
        public static Random getInstance()//单例模式
        {
            if (instance == null)
                instance = new Random();
            return instance;
        }
        public static void clear()
        {
            instance = null;
        }
        private Random() 
        {
        }
        public void rannumber(int max)
        {
            try
            {
                if (haven.Count < max - 4)
                {
                    studentnum = max;
                    while (result.Count < 4)
                    {
                        int temp = Math.Abs(random.Next(max));
                        while (haven.Contains(temp) || result.Contains(temp))
                        {
                            lock (this)
                            {
                                temp = Math.Abs(random.Next(max));
                            }
                        }
                        result.Add(temp);
                    }
                }
                else
                {
                    result.Add(-1);
                    result.Add(-1);
                    result.Add(-1);
                    result.Add(-1);
                }
            }
            catch
            {
                return;
            }
            //haven.Add(result);
        }
        public ArrayList get(int i =1)
        {
            ArrayList temp = new ArrayList();
            if (i == 1)
            {
                temp.Add((int)result[0]);
                result.RemoveAt(0);
            }
            if (i == 2)
            {
                temp.Add((int)result[0]);
                temp.Add((int)result[1]);
                result.RemoveAt(0);
                result.RemoveAt(0);
            }
            if(i==4)
            {
                temp.Add((int)result[0]);
                temp.Add((int)result[1]);
                temp.Add((int)result[2]);
                temp.Add((int)result[3]);
                result.Clear();
            }
            rannumber(studentnum);
            return temp;
        }
        public void setnumber(int number)
        {
            lock (this)
            {
                if (!haven.Contains(number))
                    haven.Add(number);
            }
        }

    }
}
