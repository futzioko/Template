using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_4335
{
	public class Usluga
	{
		public int Id { get; set; }
		public string Name { get; set; }
		public string Type { get; set; }
		public string Cost { get; set; }
	}

	public class EntityModelContainer : DbContext
	{
		public EntityModelContainer()
			: base("name=EntityModelContainer")
		{
		}

		public DbSet<Usluga> Uslugas { get; set; }
	}

}
