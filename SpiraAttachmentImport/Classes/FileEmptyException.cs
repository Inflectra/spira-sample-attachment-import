using System;

namespace SpiraAttachmentImport.Classes
{
    internal class FileEmptyException : Exception
    {
        internal FileEmptyException()
            : base("File has no contents.")
        { }
    }
}
