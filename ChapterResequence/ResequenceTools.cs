using System;
//using Microsoft.Office.Tools.Word;

using System.Collections;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
//see https://wordmvp.com/FAQs/General/UsingWildcards.htm
namespace ChapterResequence
{
    public class ResequenceTools
    {
        static public Word.Application WordApp { get; set; }
        static public Word.Document ChapterWordDoc { get; set; }

        static public int ChapterNumber { get; set; }
        static public int NewChapterNumber { get; set; }

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



        public static void ResequenceItems(string chapterWordFilePath, int chapterNumber, string itemStyle, string itemName, bool chooseDocument = false)
        {

            
            ArrayList results = GetCurrentItemOrder(chapterWordFilePath, chapterNumber, itemStyle, chooseDocumentWithWord: chooseDocument);

            foreach (string selectedListing in results)
            {
                //add ".TEMP" to all listings and references to listings
                findAndReplaceAllByText(selectedListing, selectedListing + ".TEMP");

            }

            for (int listingIndex = 1; listingIndex - 1 < results.Count; listingIndex++)
            {
                //replace all listing numbers with there proper indexes.
                findAndReplaceAllByText(results[listingIndex - 1].ToString() + ".TEMP", $"{itemName} {ChapterNumber}.{listingIndex}");

            }

            Console.WriteLine(results);

            CloseDocument(chapterWordFilePath);

            //rng.Find.get_Style();

            //rng.set_Style(ref styleNameObject1);
            //rng.Find.Forward = true;
            //rng.Find.Wrap = Word.WdFindWrap.wdFindStop;

            //change this property to true as we want to replace format
            //rng.Find.Format = true;

            //rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

        }

        public static ArrayList GetCurrentItemOrder(string chapterWordFilePath, int chapterNumber, string itemStyle, bool useFullItemTitle = false, bool chooseDocumentWithWord=false)
        {

            ChapterNumber = chapterNumber;
            if (chooseDocumentWithWord) {
                OpenDocumentViaWord();
            }
            else
            {
                OpenDocument(chapterWordFilePath);
            }

            object start = ChapterWordDoc.Content.Start;
            object end = ChapterWordDoc.Content.End;
            Word.Range rng = ChapterWordDoc.Range(ref start, ref end);



            object styleNameObject1 = itemStyle;

            ArrayList results = new ArrayList();

            rng.Find.ClearFormatting();
            rng.Find.set_Style(itemStyle);

            Word.Selection selectedListing = WordApp.Selection;

            char[] trimCharacters = { ' ', ':', '\u2002' }; // some items have weird spacing characters
            object wdLine = 5; //5==wdLine - see https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdunits?view=word-pia

            string itemTitleRegex = @"[a-zA-Z]*\ \d{1,2}\.\d{1,2}[a-zA-Z]{0,1}";


            rng.Find.Execute(); // look for first item
            while (rng.Find.Found)
            {

                selectedListing.Start = rng.Start;
                string item;
                if (useFullItemTitle) //use the text title as well. e.g. "Listing 19.5: This is a listing about..."
                {

                    selectedListing.EndKey(wdLine);//set the end of the selection to the end of the listing title line
                    selectedListing.Start = rng.Start;// set the start of selection back to the start of the range.
                    item = selectedListing.Text;
                }
                else //use only the items numbered title, else used for explicitness
                {
                    selectedListing.End = rng.End;
                    // item = selectedListing.Text.Trim(trimCharacters);
                    item = Regex.Match(selectedListing.Text, itemTitleRegex).Value;
                }

                if(!String.IsNullOrWhiteSpace(item))
                results.Add(item);

                rng.Find.ClearFormatting(); //clear the search for the newline
                rng.Find.set_Style(itemStyle);
                rng.Start = rng.Start + 100; //bodge to get the find to move forward

                rng.Find.Execute(); //look for next item
            }


            return results;
        }

        public static ArrayList AssertItemsInOrder(string chapterWordFilePath, int chapterNumber, string style)
        {

            ArrayList outofOrderItems = new ArrayList();


            ArrayList results = GetCurrentItemOrder(chapterWordFilePath, chapterNumber, style);



            int realIndex = 1;
            int index = 0;
            while (index < results.Count)
            {
                string currentItem = results[index].ToString();

                int startofItemNumber = currentItem.IndexOf($"{chapterNumber}.");
                int lengthOfChapterNum = chapterNumber.ToString().Length;
                //eg "listing 9.12" or "listing 12.1"
                // @@@ need to account for 12.12B 
                // '{chapterNumber}.' = 1 character for '.' + length of Chapter number

                string itemChapterNumberRegex = @"\d{1,2}\.";

                int itemChapterNumber = int.Parse(Regex.Match(currentItem, itemChapterNumberRegex).Value.Trim('.'));

                int nextItemLastChar;

                string nextItem = "";
                bool nextItemIsASubItem = false;
                if (index + 1 < results.Count)
                {
                    nextItem = results[index + 1].ToString();
                    nextItemIsASubItem = !int.TryParse(nextItem.Substring(nextItem.Length - 1, 1), out nextItemLastChar);
                }

                int currentItemLastChar;
                bool currentItemIsASubItem = !int.TryParse(currentItem.Substring(currentItem.Length - 1, 1), out currentItemLastChar);

                int lengthOfNumber;
                if (currentItemIsASubItem)
                {//@ if last character is a letter like in 12.5B don't increment realIndex and don't use it calculate the decimal places of the item number
                    //realIndex--; use for not incrementing, but this doesn't work when the item is out of order. like 19.15b comes after 19.16
                    lengthOfNumber = currentItem.Length - (startofItemNumber + 1 + lengthOfChapterNum + 1);
                }
                else
                {
                    lengthOfNumber = currentItem.Length - (startofItemNumber + 1 + lengthOfChapterNum);
                }
                string itemNumberRegex = @"\.\d{1,2}";

                int itemNumber = int.Parse(Regex.Match(currentItem, itemNumberRegex).Value.Trim('.'));

                if (itemNumber != realIndex || chapterNumber != itemChapterNumber)
                {
                    outofOrderItems.Add($"itemNumber {currentItem} should be {chapterNumber}.{realIndex}");
                }

                realIndex++;
                index++;
            }


            return outofOrderItems;
        }

        public static void findAndReplaceAllByText(string text, string replacementText)
        {
            //if (text.Equals("Listing 20.1")) return; // uncomment for bodge when find and replace tool doesn't work -> don't want a search for 20.1 to return 20.10,20.11.. etc

           
            Word.Range rng = WordApp.ActiveDocument.Content;

            rng.Find.ClearFormatting();

            rng.Find.Text = text; //uncomment for regular find and replace
            //rng.Find.Text = text + "^?"; //uncomment for bodge when find and replace tool not working
            rng.Find.Replacement.Text = replacementText;
            rng.Find.Forward = true;
            rng.Find.MatchCase = false;

            rng.Find.MatchSuffix = true; // so "listing 9.1" != "listing 9.17" //uncomment for regular find and replace
            
            

            //regular method
            rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll); // uncomment for regular method

            // for listing mentions like "In Lisiting 20.1's use of..." -> unfortunately Word find doesn't support zero or more wild cards so can't do a "'s" search all at once
            rng.Find.Text = text + "\'s";
            rng.Find.Replacement.Text = replacementText+"\'s";
            rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);



            /*      //bodge for not being able to find with the replace tool
                  rng.Find.Execute(); // look for first match
                  while (rng.Find.Found)
                  {
                      rng.End = rng.End - 1;

                      rng.Text = replacementText;

                      rng.Start = rng.End;

                      rng.Find.Execute(); //look for next match
                  }*/


        }



    }
}
