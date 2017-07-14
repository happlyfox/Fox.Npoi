using Fox.Npoi.Test.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Fox.Npoi.Test
{
    public class DataUtils
    {
        public static List<Depart> GetDepartList()
        {
            return Enumerable.Range(0, 10).Select(index => new Depart()
            {
                DepId = index,
                DepName = string.Format("测试部门", index)
            }).ToList();
        }

        public static List<User> GetUserList()
        {
            List<Depart> depList = Enumerable.Range(0, 10).Select(index => new Depart()
            {
                DepId = index,
                DepName = string.Format("测试部门", index)
            }).ToList();

            List<User> userList = new List<User>();
            Random random = new Random();
            for (int i = 0; i < 100; i++)
            {
                int depid = random.Next(0, depList.Count);
                userList.Add(new User
                {
                    Id = i,
                    Name = string.Format("测试人员{0}", i),
                    Age = 10,
                    Address = string.Format("测试地址{0}", i),
                    DepId = depid
                });
            }
            return userList;
        }
    }
}