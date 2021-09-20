using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ChapterResequenceTest
{
    /// <summary>
    /// Use these tests to run individual programs.
    /// </summary>
    [TestClass]
    public class TestsThatRunIndividualTools
    {
        [TestMethod]
        public void ParseNumFromFileName_Success()
        {
            Assert.IsTrue(ChapterResequence.Program.ParseChapterNumber("Michaelis_Ch09") == 9);

            Assert.IsTrue(ChapterResequence.Program.ParseChapterNumber("Michaelis_Ch12") == 12);

        }

        /*[TestMethod]
        public void OrderOfListings()
        {

            string testFile = @"..\..\Michaelis_Ch12 - full.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));
            
            var x = ChapterResequence.ResequenceTools.AssertItemsInOrder(Path.GetFullPath(testFile), 12, "CDT_NUM");

            foreach (string t in x)
            {
                Trace.WriteLine(t);
            }

            ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }*/

        [TestMethod]
        public void GetFullTitleOrderOfListings()
        {

            string testFile = @"..\..\Michaelis_Ch12 - full.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            string[] listingTitleStyle = { " " };

            var x = ChapterResequence.ResequenceTools.GetCurrentItemOrder(Path.GetFullPath(testFile), 12, "CDT_NUM", true);

            foreach (string t in x)
            {
                Trace.WriteLine(t);
            }

            ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }

        [TestMethod]
        public void GetOrderOfListings()
        {
            //"C:\Users\saffron\Desktop\TextBookTools\ChapterResequenceTest\Michaelis_Ch22 (formerly 20).docx"
            string testFile = @"..\..\Michaelis_Ch18.docx";
            int chapterNumber = 18;
            //string testFile = @"..\..\Michaelis_Ch12.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            string cdtNum = "CDT_NUM";
            string listingNumber = "ListingNumber";
            //different chapters have different styles;
            string listingStyle = listingNumber;

            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            var x = ChapterResequence.ResequenceTools.GetCurrentItemOrder(Path.GetFullPath(testFile), chapterNumber, cdtNum, useFullItemTitle: true);

            foreach (string t in x)
            {
                Trace.WriteLine(t);
            }

            ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }


        [TestMethod]
        public void AssertOrderOfListings()
        {
            //"C:\Users\saffron\Desktop\TextBookTools\ChapterResequenceTest\Michaelis_Ch22 (formerly 20).docx"
            string testFile = @"..\..\Michaelis_Ch18.docx";
            int chapterNumber = 18;
            //string testFile = @"..\..\Michaelis_Ch12.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            string cdtNum = "CDT_NUM";
            string listingNumber = "ListingNumber";
            //different chapters have different styles;
            string listingStyle = listingNumber;


            var x = ChapterResequence.ResequenceTools.AssertItemsInOrder(Path.GetFullPath(testFile), chapterNumber, listingStyle);

            foreach (string t in x)
            {
                Trace.WriteLine(t);
            }

            ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }


        [TestMethod]
        public void ResequenceListings()
        {
            //"C:\Users\saffron\Desktop\TextBookTools\ChapterResequenceTest\Michaelis_Ch22 (formerly 20).docx"
            string testFile = @"..\..\Michaelis_Ch20.docx"; // if chooseDocument set to false
            int chapterNumber = 8;
          
            //Trace.WriteLine(Directory.GetCurrentDirectory());
            //Assert.IsTrue(File.Exists(testFile));

            string cdtNum = "CDT_NUM";
            string listingNumber = "ListingNumber";
            //different chapters have different styles;



            ChapterResequence.ResequenceTools.ResequenceItems(Path.GetFullPath(testFile), chapterNumber, cdtNum, "Listing", chooseDocument: true);

            //ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }

        [TestMethod]
        public void ResequenceOutputs()
        {
            //"C:\Users\saffron\Desktop\TextBookTools\ChapterResequenceTest\Michaelis_Ch22 (formerly 20).docx"
            string testFile = @"..\..\Michaelis_Ch22 (formerly 20).docx";
            int chapterNumber = 23;
            //string testFile = @"..\..\Michaelis_Ch12.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));



            //different chapters have different styles;
            string style = "OutputNumber";


            ChapterResequence.ResequenceTools.ResequenceItems(Path.GetFullPath(testFile), chapterNumber, style, "Output", chooseDocument: true);

            //ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }


        [TestMethod]
        public void ResequenceTables()
        {
            //"C:\Users\saffron\Desktop\TextBookTools\ChapterResequenceTest\Michaelis_Ch22 (formerly 20).docx"
            string testFile = @"..\..\Michaelis_Ch22 (formerly 20).docx";
            int chapterNumber = 23;
            //string testFile = @"..\..\Michaelis_Ch12.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));



            //different chapters have different styles;
            string style = "TableNumber";


            ChapterResequence.ResequenceTools.ResequenceItems(Path.GetFullPath(testFile), chapterNumber, style, "Table", chooseDocument: true);

            //ChapterResequence.ResequenceTools.CloseDocument(Path.GetFullPath(testFile));

        }

        [TestMethod]
        public void GetKeyWords() //this will append all new unique keywords to the keywords.txt from the chosen document
        {

            string testFile = @"..\..\Michaelis_Ch19.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            var x = ChapterResequence.KeyWordTools.GetCSharpKeyWordsInChapter(chapterWordFilePath: Path.GetFullPath(testFile), true);

            foreach (string t in x)
            {
                Trace.WriteLine(t);
            }


        }

        [TestMethod]
        public void SetKeyWordStyle()
        {

            string testFile = @"..\..\Michaelis_Ch19.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            string listingTitle = "Listing 5.12";
            string listingTitleStyle = "CDT_NUM";
            string startOfListingStyle = "CDT_FIRST";
            string endOfListingStyle = "CDT_LAST";

            // ChapterResequence.KeyWordTools.StylizeKeyWords(listingTitle, chapterWordFilePath: Path.GetFullPath(testFile));

            ChapterResequence.KeyWordTools.StylizeKeyWords(listingTitle, listingTitleStyle, startOfListingStyle, endOfListingStyle, chooseDocumentWithWord: true);

            /* foreach (string t in x)
             {
                 Trace.WriteLine(t);
             }*/

        }


        [TestMethod]
        public void GetListingRange()
        {

            string testFile = @"..\..\Michaelis_Ch19.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));

            string listingTitle = "Listing 20.2";
            string listingTitleStyle = "CDT_NUM";
            string startOfListingStyle = "CDT_FIRST";
            string endOfListingStyle = "CDT_LAST";

            ChapterResequence.KeyWordTools.OpenDocument(Path.GetFullPath(testFile));

            Word.Range _ = ChapterResequence.KeyWordTools.GetListingRange(listingTitle, listingTitleStyle, startOfListingStyle, endOfListingStyle);


        }


        [TestMethod]
        public void GetGuidelines()
        {

            string testFile = @"..\..\Michaelis_Ch19.docx";
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Assert.IsTrue(File.Exists(testFile));



            ChapterResequence.GuideLineTools.GetGuideLinesInDocument(Path.GetFullPath(testFile));

            /* foreach (string t in x)
             {
                 Trace.WriteLine(t);
             }*/



        }

    }
}
