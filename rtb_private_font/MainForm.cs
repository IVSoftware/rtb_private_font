using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static rtb_private_font.FontNameExtractor;

// https://superuser.com/questions/120593/how-do-you-change-a-ttf-font-name
// https://social.msdn.microsoft.com/Forums/vstudio/en-US/77e90934-1096-425c-bfad-981769c3d930/unable-to-change-the-selectionfontname-property-in-richtextbox?forum=winforms
// http://www.biblioscape.com/rtf15_spec.htm#:~:text=RTF%20supports%20embedded%20fonts%20with,contained%20in%20the%20%5Cfontfile%20group.
// https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getfontdata
// http://pinvoke.net/default.aspx/gdi32/GetFontData.html
// http://pinvoke.net/default.aspx/gdi32/SelectObject.html
// https://www.pinvoke.net/default.aspx/gdi32/DeleteObject.html

namespace rtb_private_font
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Fonts", "smiley.ttf");
            privateFontCollection.AddFontFile(path);

            var fontFamily = privateFontCollection.Families[0];
            Debug.Assert(fontFamily.Name == "smiley", "Expecting 'smiley' is the font family name");

            var font = new Font(fontFamily, 12F);

            richTextBox1.Text = "\uE800";
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.Select(0, 1);
            richTextBox1.SelectionFont = font;

            // WORKAROUND: You can obtain the result using the OriginalFontName property
            var privateFontName = font.Name;
            var selectionFontName = richTextBox1.SelectionFont.Name;
            var selectionOriginalFontName = richTextBox1.SelectionFont.OriginalFontName;

            richTextBox1.SelectionBackColor = Color.Yellow;
            richTextBox1.AppendText($"PrivateFont={privateFontName}{Environment.NewLine}");
            richTextBox1.AppendText($"SelectionFont.Name={selectionFontName}{Environment.NewLine}");
            richTextBox1.AppendText($"SelectionFont.OriginalFontName={selectionOriginalFontName}{Environment.NewLine}");

            var rtf = richTextBox1.Rtf;
            { }

            richTextBox1.Select(int.MaxValue, 0);

            // This gets the same result for SelectionFont
            richTextBox1.SelectionChanged += RichTextBox1_SelectionChanged;

            // The private font is read back correctly when
            // used for the Font property of a RichTextBox
            richTextBox2.Font = font;
            richTextBox2.AppendText($"Font={richTextBox2.Font.Name}");
            richTextBox2.SelectAll();
            richTextBox2.SelectionFont = richTextBox1.Font;
            richTextBox2.Select(int.MaxValue, 0);

            richTextBox1.SaveFile("output.rtf");

            // The font can be auto-installed like this.
            if(IsFontInstalled("smiley"))
            {
                // Test to see if private font is available in the doc
                Process.Start("explorer.exe", "output.rtf");
            }
            else
            {
                Process.Start("explorer.exe", path);
            }
        }
        // https://stackoverflow.com/a/53006947/5438626
        static bool IsFontInstalled(string fontname)
        {
            using (var ifc = new InstalledFontCollection())
            {
                return ifc.Families.Any(f => f.Name == fontname);
            }
        }

        PrivateFontCollection privateFontCollection = new PrivateFontCollection();

        private void RichTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            var selectionFont = richTextBox1.SelectionFont;
            if (richTextBox1.SelectedText.Length == 1)
            {
                richTextBox1.AppendText($"SelectionFont={richTextBox1.SelectionFont.OriginalFontName}{Environment.NewLine}");
            }
        }
    }
}
