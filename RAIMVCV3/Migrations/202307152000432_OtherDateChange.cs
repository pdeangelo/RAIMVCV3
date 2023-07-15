namespace RAIMVCV3.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class OtherDateChange : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.Loans", "LoanFundingDate", c => c.DateTime());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.Loans", "LoanFundingDate", c => c.DateTime(storeType: "date"));
        }
    }
}
