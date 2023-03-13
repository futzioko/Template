using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_4335.Models
{
    internal class UserContext : DbContext
    {
        public UserContext() : base("ISRPO_LR2") { }

        public DbSet<User> Users { get; set; }
        public DbSet<Streets> Streets { get; set; }
    }
}
