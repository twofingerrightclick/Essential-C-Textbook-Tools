using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace ChapterResequence
{
  
    class Utility
    {
   
        public static void OpenDocument(String docPath, out Word.Application wordApp, out Word.Document wordDoc )
        {
            // Get the Word application object.
            wordApp = new Word.Application();

            // Make Word visible (optional).
            wordApp.Visible = true;

            wordDoc = wordApp.Documents.Open(docPath);

            wordDoc.TrackRevisions = true;

            wordApp.Activate();
        }

        public static void OpenDocumentViaWord(out Word.Application wordApp, out Word.Document wordDoc)
        {
            // Get the Word application object.
            wordApp = new Word.Application();

            // Make Word visible (optional).
            wordApp.Visible = true;

            wordApp.Activate();



            wordDoc = wordApp.ActiveDocument;

            wordDoc.TrackRevisions = true;


        }

        public static void CloseDocumentNoSave(Word.Application wordApp)
        {
            wordApp.Quit(SaveChanges:Word.WdSaveOptions.wdDoNotSaveChanges);
        }

        public static void CloseDocument(Word.Application wordApp)
        {
            wordApp.Quit();
        }

    }
}
