using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpiraAttachmentImport.Classes
{
    /// <summary>
    /// Represents an artifact mapping
    /// </summary>
    public class ArtifactMapping
    {
        /// <summary>
        /// The filename of the attachment
        /// </summary>
        public string Filename { get; set; }

        /// <summary>
        /// The type of Spira artifact being attached to
        /// </summary>
        public int ArtifactTypeId { get; set; }

        /// <summary>
        /// The identifier of the artifact, it will be stored as a custom property value on the artifact
        /// </summary>
        public string ExternalKey { get; set; }
    }
}
