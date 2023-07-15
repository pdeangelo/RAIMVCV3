namespace RAIMVCV3.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class OtherDatesDummyChange : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.Loans", "LoanEnteredDate", c => c.DateTime());
            AlterColumn("dbo.Loans", "DateDepositedInEscrow", c => c.DateTime());
            AlterColumn("dbo.Loans", "BaileeLetterDate", c => c.DateTime());
            AlterColumn("dbo.Loans", "ClosingDate", c => c.DateTime());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.Loans", "ClosingDate", c => c.DateTime(storeType: "date"));
            AlterColumn("dbo.Loans", "BaileeLetterDate", c => c.DateTime(storeType: "date"));
            AlterColumn("dbo.Loans", "DateDepositedInEscrow", c => c.DateTime(storeType: "date"));
            AlterColumn("dbo.Loans", "LoanEnteredDate", c => c.DateTime(storeType: "date"));
        }
    }
}
