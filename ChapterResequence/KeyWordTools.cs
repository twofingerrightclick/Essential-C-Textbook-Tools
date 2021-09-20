using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace ChapterResequence
{
    public static class KeyWordTools
    {

        static public Word.Application WordApp { get; set; }
        static public Word.Document ChapterWordDoc { get; set; }

        static string _KeywordFilePath = @"C:\Users\saffron\Desktop\TextBookTools\ChapterResequence\keywords.txt";

        static public int ChapterNumber { get; set; }
        public static void OpenDocument(String docPath)
        {
            // Get the Word application object.
            WordApp = new Word.Application();

            // Make Word visible (optional).
            WordApp.Visible = true;

            ChapterWordDoc = WordApp.Documents.Open(docPath);

            ChapterWordDoc.TrackRevisions = true;

            WordApp.Activate();
        }

        public static void OpenDocumentViaWord()
        {
            // Get the Word application object.
            WordApp = new Word.Application();

            // Make Word visible (optional).
            WordApp.Visible = true;

            WordApp.Activate();



            ChapterWordDoc = WordApp.ActiveDocument;

            ChapterWordDoc.TrackRevisions = true;


        }

        public static void CloseDocument(String chapterWordFilePath)
        {
            WordApp.Quit();
        }


        public static Word.Range GetListingRange(string listingTitle, string listingTitleStyle, string startOfListingStyle, string endOfListingStyle)
        {

            object start = ChapterWordDoc.Content.Start;
            object end = ChapterWordDoc.Content.End;
            Word.Range rng = ChapterWordDoc.Range(ref start, ref end);
            string listingStyle = listingTitleStyle ?? "ListingNumber";
            rng.Find.Text = listingTitle;
            rng.Find.set_Style(listingStyle);
            rng.Find.Execute();
            int startOfListingIndex = rng.End;

            endOfListingStyle = endOfListingStyle ?? "CDTX";

            object startRange = rng.End;
            object endRange = ChapterWordDoc.Content.End;
            Word.Range endOfListingRange = ChapterWordDoc.Range(ref startRange, ref endRange);

            endOfListingRange.Find.ClearFormatting();
            endOfListingRange.Find.set_Style(endOfListingStyle);
            endOfListingRange.Find.Execute();
            int endOfListingIndex = endOfListingRange.End;

            object startListing = startOfListingIndex;
            object endListing = endOfListingIndex;
            Word.Range listingRange = ChapterWordDoc.Range(ref startListing, ref endListing);

            return listingRange;

        }

        //C# keywords in listings
        public static void StylizeKeyWords(string listingTitle, string listingTitleStyle, string startOfListingStyle, string endOfListingStyle, bool chooseDocumentWithWord = false, string chapterWordFilePath = null)
        {

            //HashSet<string> keyWords = GetCSharpKeyWordsInChapter(chapterWordFilePath, chooseDocumentWithWord);
            if (chooseDocumentWithWord)
            {
                OpenDocumentViaWord();
            }
            else
            {
                OpenDocument(chapterWordFilePath);
            }
            HashSet<string> keyWords = GetCSharpKeyWordsFromKeyWordFile();

            Word.Range listingRange = GetListingRange(listingTitle, listingTitleStyle, startOfListingStyle, endOfListingStyle);

            string keyWordStyle = "CP Keyword";

            object keyWordStyleObject = keyWordStyle;

            string[] listingWords = listingRange.Text.Split();
            HashSet<string> uniqueWordsinListing = new HashSet<string>();

            foreach (string s in listingWords)
            {
                s.Trim();
                if (keyWords.Contains(s))
                {
                    uniqueWordsinListing.Add(s);
                }
            }

            int listingRangeStart = listingRange.Start;
            int listingRangeEnd = listingRange.End;
            object start = listingRangeStart;
            object end = listingRangeEnd;

            //foreach keyword search through the listing
            foreach (string keyWord in uniqueWordsinListing)
            {
                
                Word.Range subRange = ChapterWordDoc.Range(ref start, ref end);
                subRange.Find.ClearFormatting();

                subRange.Find.Text = keyWord;

                subRange.Find.MatchWholeWord = true; //only change independent words


                subRange.Find.Execute(); // look for first instance of keyword
                while (subRange.Find.Found)
                {
                    if (subRange.End > listingRangeEnd) break;

                    subRange.set_Style(keyWordStyleObject);

                    subRange.Find.Execute(); //look for other instances
                }


            }

            FormatStrings(start, end, listingRangeStart,listingRangeEnd) ;

        }

        private static void FormatStrings(object start, object end, int listingRangeStart, int listingRangeEnd)
        {
           
            Word.Range subRange = ChapterWordDoc.Range(ref start, ref end);
            subRange.Find.ClearFormatting();
            subRange.Find.MatchWildcards = true;
            subRange.Find.Text = "\"*\"";

            string stringStyle = "Maroon";

            object stringStyleObject = stringStyle;


            subRange.Find.Execute(); // look for first instance of keyword
            while (subRange.Find.Found)
            {
                if (subRange.End > listingRangeEnd) break;

                subRange.set_Style(stringStyleObject);

                //subRange.Find.Execute(); //the first execute finds the last " and connects it with the next "
                subRange.Find.Execute();// the second execute gets the actual thing we wanted between quotes
            }


            //format special characters
            subRange = ChapterWordDoc.Range(ref start, ref end);
            subRange.Find.ClearFormatting();
            subRange.Find.MatchWildcards = true;
            subRange.Find.Text = "[$\\@]\"";

            subRange.Find.Execute(); // look for first instance of keyword
            while (subRange.Find.Found)
            {
                if (subRange.End > listingRangeEnd) break;

                subRange.set_Style(stringStyleObject);

               
                subRange.Find.Execute();
            }

            //undo formatting of string interpolations-->
            //string midListingStyle = "CDT_MID";
            //object midListingStyleObject = midListingStyle;
            subRange = ChapterWordDoc.Range(ref start, ref end);
            subRange.Find.ClearFormatting();
            subRange.Find.MatchWildcards = true;
            subRange.Find.set_Style(stringStyleObject);
            subRange.Find.Text = "\\{*\\}";

            subRange.Find.Execute(); 
            while (subRange.Find.Found)
            {
                if (subRange.End > listingRangeEnd) break;

                subRange.Font.Color = Word.WdColor.wdColorBlack;

                subRange.Find.Execute();
            }
        }

        public static HashSet<string> GetCSharpKeyWordsInChapter(string chapterWordFilePath, bool usewordtoOpenDocument)

        {

            if (usewordtoOpenDocument)
            {
                OpenDocumentViaWord();
            }
            else
            {
                OpenDocument(chapterWordFilePath);
            }

            object start = ChapterWordDoc.Content.Start;
            object end = ChapterWordDoc.Content.End;
            Word.Range rng = ChapterWordDoc.Range(ref start, ref end);

            string keyWordStyle = "CP Keyword";

            object styleNameObject1 = keyWordStyle;

            HashSet<string> results = new HashSet<string>();

            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();
            rng.Find.Forward = true;

            rng.Find.set_Style(keyWordStyle);

            Word.Selection selectedKeyword = WordApp.Selection;

            rng.Find.Execute(); // look for first keyword
            while (rng.Find.Found)
            {

                selectedKeyword.Start = rng.Start;
                int startx = rng.Start;
                int endx = rng.End;
                selectedKeyword.End = rng.End;

                char[] trimCharacters = { ' ' };

                string keywords = selectedKeyword.Text.Trim(trimCharacters);
                string[] keyWordsSplit = keywords.Split(trimCharacters);
                foreach (string item in keyWordsSplit)
                {
                    string keyword = Regex.Replace(item, @"\t|\n|\r", "");
                    if (keyword.Length > 1 && !String.IsNullOrWhiteSpace(item))
                    {
                        results.Add(keyword);
                    }
                }
                rng.Start = rng.Start + 10; //bodge to get the rng Find to move forward in dumb documents!
                rng.Find.Execute(); //look for next keyword
            }

            return WriteNewKeywordsToFile(results); //checks current files keywords with keywords file and returns all unique keywords from those sources
            
        }

        public static HashSet<string> WriteNewKeywordsToFile(HashSet<string> keywords) {

            HashSet<string> currentKeyWordsInFile = GetCSharpKeyWordsFromKeyWordFile();
            foreach (string s in currentKeyWordsInFile )
            {
                keywords.Add(s);
            }

           using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(_KeywordFilePath))
            {
                foreach (string keyword in keywords)
                {

                    file.WriteLine(keyword);       
                }
            }

            return keywords;

        }

        private static HashSet<string> GetCSharpKeyWordsFromKeyWordFile()
        {
            HashSet<string> keywords = new HashSet<string>();
            

            // This text is added only once to the file.
            if (!File.Exists(_KeywordFilePath))
            {
                // Create a file to write to.
                File.Create(_KeywordFilePath);
            }

            // check for existing keywords.
            string[] readText = File.ReadAllLines(_KeywordFilePath);
            foreach (string s in readText)
            {
                keywords.Add(s);
            }

            return keywords;
        }




    }

     
}
