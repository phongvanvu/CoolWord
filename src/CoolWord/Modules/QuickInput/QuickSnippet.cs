using System;
using System.Runtime.Serialization;

namespace CoolWord.Modules.QuickInput
{
    [DataContract]
    public class QuickSnippet
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string Title { get; set; }

        [DataMember]
        public string Category { get; set; }

        [DataMember]
        public string Content { get; set; }

        [DataMember]
        public string Shortcut { get; set; }

        [DataMember]
        public bool IsBuiltIn { get; set; }

        public QuickSnippet()
        {
            Id = Guid.NewGuid().ToString("N");
            Title = string.Empty;
            Category = "Chung";
            Content = string.Empty;
            Shortcut = string.Empty;
            IsBuiltIn = false;
        }

        public QuickSnippet(string id, string title, string category, string content, string shortcut = null, bool isBuiltIn = false)
        {
            Id = id ?? Guid.NewGuid().ToString("N");
            Title = title ?? string.Empty;
            Category = category ?? "Chung";
            Content = content ?? string.Empty;
            Shortcut = shortcut ?? string.Empty;
            IsBuiltIn = isBuiltIn;
        }

        public override string ToString()
        {
            return string.Format("[{0}] {1}", Category, Title);
        }
    }
}
