using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MsftGraphBotQuickStartLUIS.Helpers
{
    public class QueryBuilder
    {
        public static string GetFileNameQuery(string queryTemplate, string entityValue)
        {
            return String.Format(queryTemplate, entityValue.Replace(" . ", ".").Replace(". ", ".")); ;
        }

        public static string GetFileTypeQuery(string queryTemplate, string entityValue)
        {
            // Building the file types for the OneDrive search based on the entity specified
            var fileType = entityValue.Replace(" . ", ".").Replace(". ", ".").ToLower();
            List<string> images = new List<string>() { "images", "pictures", "pics", "photos", "image", "picture", "pic", "photo" };
            List<string> presentations = new List<string>() { "powerpoints", "presentations", "decks", "slidedecks", "powerpoints", "presentation", "deck" };
            List<string> documents = new List<string>() { "documents", "document", "word" };
            List<string> workbooks = new List<string>() { "workbooks", "workbook", "excel", "spreadsheet", "spreadsheets" };
            List<string> music = new List<string>() { "music", "songs", "albums", "tunes" };
            List<string> videos = new List<string>() { "video", "videos", "movie", "movies" };

            String query = String.Empty;

            // Building the search query based on the filetype value from the entity
            if (images.Contains(fileType))
                query = String.Format(queryTemplate, ".png OR .jpg OR .jpeg OR .gif");
            else if (presentations.Contains(fileType))
                query = String.Format(queryTemplate, ".pptx OR .ppt");
            else if (documents.Contains(fileType))
                query = String.Format(queryTemplate, ".docx OR .doc");
            else if (workbooks.Contains(fileType))
                query = String.Format(queryTemplate, ".xlsx OR .xls");
            else if (music.Contains(fileType))
                query = String.Format(queryTemplate, ".mp3 OR .wav");
            else if (videos.Contains(fileType))
                query = String.Format(queryTemplate, ".mp4 OR .avi OR .mov");
            else
                query = String.Format(queryTemplate, fileType);

            // Returns the formatted query
            return query;
        }
    }
}