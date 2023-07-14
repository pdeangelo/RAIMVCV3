namespace RAIMVCV3.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddIsSelected : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.Loans", "IsChecked", c => c.Boolean(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.Loans", "IsChecked");
        }
    }
}
