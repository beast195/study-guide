namespace SudyGuide.Framework.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class sentenceLibraryTbl : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.SentenceLibraries",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        VerbsPart = c.String(),
                        AdjectitivesPart = c.String(),
                        PrepositionPart = c.String(),
                        AdverbsPart = c.String(),
                        NounsPart = c.String(),
                        PronounsPart = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.SentenceLibraries");
        }
    }
}
