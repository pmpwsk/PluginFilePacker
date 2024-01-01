using Microsoft.VisualStudio.Shell;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace uwap.VSIX.PluginFilePacker
{
    internal class Options : DialogPage
    {
        private bool _UseBase64 = true;
        [Category("PluginFilePacker")]
        [DisplayName("UseBase64")]
        [Description("Whether to use base64 to encode files (true, slightly more efficient runtime) or to specify files as byte arrays (false, way slower builds).")]
        public bool UseBase64
        {
            get { return _UseBase64; }
            set { _UseBase64 = value; }
        }

        private string _TextExtensions = "css,js,txt,json";
        [Category("PluginFilePacker")]
        [DisplayName("TextExtensions")]
        [Description("The file extensions (without the preceding dot) that should be recognized as text files, separated by a comma/semicolon/space.")]
        public string TextExtensions
        {
            get { return _TextExtensions; }
            set { _TextExtensions = value; }
        }
        internal IEnumerable<string> TextExtensionsClean
            => _TextExtensions.Split(',', ';', ' ').Select(x => x.Trim()).Where(x => x != "").Select(x => x[0] == '.' ? x : ('.' + x));

        private string _DefaultNamespace = "uwap.WebFramework.Plugins";
        [Category("PluginFilePacker")]
        [DisplayName("DefaultNamespace")]
        [Description("The default namespace to use if no existing FileHandler.cs or FileHandlerCustom.cs with a namespace was found.")]
        public string DefaultNamespace
        {
            get { return _DefaultNamespace; }
            set { _DefaultNamespace = value; }
        }

        private bool _ShowPopupWhenDone;
        [Category("PluginFilePacker")]
        [DisplayName("ShowPopupWhenDone")]
        [Description("Whether to show a popup once FileHandler.cs was successfully generated.")]
        public bool ShowPopupWhenDone
        {
            get => _ShowPopupWhenDone;
            set => _ShowPopupWhenDone = value;
        }
    }
}
