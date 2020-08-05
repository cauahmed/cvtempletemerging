using Aspose.Words;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Aspose.Words.Replacing;
using Aspose.Words.Settings;
using System.Text.RegularExpressions;

namespace TestApp2
{
    class Program
    {
        static void InsertDocument(Node insertAfterNode, Document sourceDoc)
        {
            //Check if the node is a paragraph or a table
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) & (!insertAfterNode.NodeType.Equals(NodeType.Table)))
            {
                throw new ArgumentException("The destination node should either be a paragraph or a table");
            }

            //Insert into the parent of the destination paragraph
            CompositeNode destinationStory = insertAfterNode.ParentNode;

            //Translating styles and lists during the import
            NodeImporter importer = new NodeImporter(sourceDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            //Loop through all sections of the source document
            foreach (Section sourceSection in sourceDoc.Sections)
            {
                //Loop through all block level nodes such as paragraphs and tables in the body section
                foreach (Node sourceNode in sourceSection.Body)
                {
                    //skip the node if its the last empty pargraph of the section
                    if (sourceNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Paragraph para = (Paragraph)sourceNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                        {
                            continue;
                        }
                    }

                    //creating a clone of the node to be inserted in the destination document
                    Node newNode = importer.ImportNode(sourceNode, true);

                    //Insert the clone(new node) after the reference node
                    destinationStory.InsertAfter(newNode, insertAfterNode);
                    //move on to the next node by changing reference node
                    insertAfterNode = newNode;
                }
            }

        }
        static void Main(string[] args)
        {
            const string dataDir = "C:\\Personal Development\\CV merging\\Data\\";
            Document DestinationDoc = new Document(dataDir + "InsertDocument1.doc");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertDocumentReplaceHandler();

            DestinationDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"), "", options);
            DestinationDoc.Save(dataDir + "InsertDocumentAtReplace_out.doc");
        }

        private class InsertDocumentReplaceHandler : IReplacingCallback
        {
            

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                Document subDoc = new Document("C:\\Personal Development\\CV merging\\Data\\CV Out.docx");

                //Insert a document after the paragraph, containing the matched text
                Paragraph para = (Paragraph)args.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                //Remove the paragraph with the matched text
                para.Remove();

                return ReplaceAction.Skip;
            }
        }
    }
}
