using System.Data.Entity;

namespace AZSWeb.Models
{
    public partial class User
    {
        public int ID { get; set; } = -1;
        public string Name { get; set; } = "";
        public string Pass { get; set; } = "";

        public User()
        {

        }
    }

    public class UserContext : DbContext
    {
        public UserContext()
            : base("DbConnection")
        { }

        public DbSet<User> Users { get; set; }
    }


}