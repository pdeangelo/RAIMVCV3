namespace RAIMVCV3.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InvestorproceedsDummyChange : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.Loans", "InvestorProceedsDate", c => c.DateTime());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.Loans", "InvestorProceedsDate", c => c.DateTime(storeType: "date"));
        }
    }
}
