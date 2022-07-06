using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text; 
// http://pinvoke.net/default.aspx/gdi32/GetFontData.html

namespace rtb_private_font
{
    /// <summary>
    /// Helper to enable loading the full font name as specified in TrueType font specs.
    /// http://www.microsoft.com/typography/otspec/name.htm
    /// </summary>
    internal static class FontNameExtractor
    {
        #region PInvoke gdi32.dll

        /// <summary>
        /// Selects the graphics object (here a font handle) into the device context.
        /// </summary>
        /// <param name="hdc">A handle to the device context.</param>
        /// <param name="hgdiobj">A handle to the object to be selected.
        ///  The specified object must have been created by using one of the following functions.</param>
        /// <returns>If the selected object is not a region and the function succeeds, the return value is a handle to the object being replaced.
        /// If an error occurs and the selected object is not a region, the return value is NULL. Otherwise, it is HGDI_ERROR.</returns>
        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        /// <summary>
        /// Deletes a logical pen, brush, font, bitmap, region, or palette handle freeing all system resources associated with the object.
        /// After the object is deleted, the specified handle is no longer valid.
        /// </summary>
        /// <param name="hgdiobj">The handle to the logical GDI object.</param>
        /// <returns><c>true</c>, if the handle to the GDI object was successfully released; otherwise <c>false</c> is returned.</returns>
        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hgdiobj);

        /// <summary>
        /// Gets the font data from specified font metric table to be accessed by key defined as <paramref name="dwTable"/>.
        /// </summary>
        /// <param name="hdc">The device context, the font file was loaded into.
        ///  Use <see cref="SelectObject"/> for loading the font into a device context.</param>
        /// <param name="dwTable">The key of the font metric table as <see cref="uint"/>
        /// assembled by ASCII HEX code of the single characters of the table in reversed order.
        /// <code>uint table = 0x656D616E</code>, represents the font table keyword character sequence 'name' in reversed order (i.e. 'eman').</param>
        /// <param name="dwOffset">Specifies the dwOffset from the beginning of the font metric table to the location where the function should begin retrieving information.
        /// If this parameter is zero, the information is retrieved starting at the beginning of the table specified by the <paramref name="dwTable"/> parameter.
        /// If this value is greater than or equal to the size of the table, an error occurs.</param>
        /// <param name="lpvBuffer">Points to a lpvBuffer to receive the font information. If this parameter is NULL, the function returns the size of the <paramref name="lpvBuffer"/> required for the font data.</param>
        /// <param name="cbData">Specifies the length, in bytes, of the information to be retrieved.
        /// If this parameter is zero, GetFontData returns the size of the font metrics table specified in the <paramref name="dwTable"/> parameter. </param>
        /// <returns>If the function succeeds, the return value is the number of bytes returned. If the function fails, the return value is GDI_ERROR.</returns>
        [DllImport("gdi32.dll")]
        private static extern uint GetFontData(IntPtr hdc, uint dwTable, uint dwOffset, [Out] byte[] lpvBuffer, uint cbData);

        #endregion

        private static uint nameTableKey = 0x656D616E;   // Represents the ASCII character sequence 'eman' => reversed name table key 'name'

        /// <summary>
        /// Gets the full name directly out of specified True Type font file.
        /// </summary>
        /// <param name="fontFile">The True Type font file.</param>
        /// <returns>The full font name.</returns>
        internal static string GetFullFontName(string fontFile)
        {
            string fontName = string.Empty;

            using (PrivateFontCollection fontCollection = new PrivateFontCollection())
            {
                fontCollection.AddFontFile(fontFile);
                byte[] fontData = LoadFontMetricsNameTable(fontCollection);

                fontName = ExtractFullName(fontData);
            }
            return fontName;
        }


        /// <summary>
        /// Gets the font data
        /// </summary>
        internal static byte[] GetFontData(string fontFile)
        {
            string fontName = string.Empty;
            using (PrivateFontCollection fontCollection = new PrivateFontCollection())
            {
                fontCollection.AddFontFile(fontFile);
                return LoadFontMetricsNameTable(fontCollection);
            }
        }

        #region Private Helper

        /// <summary>
        /// Extracts the full font name from raw bytes.
        /// </summary>
        /// <param name="fontData">The font data as raw bytes.</param>
        /// <returns>The extracted full font name.</returns>
        private static string ExtractFullName(byte[] fontData)
        {
            string fontName = string.Empty;

            using (BinaryReader br = new BinaryReader(new MemoryStream(fontData)))
            {
                // Read selector (always = 0) to advance reader position by 2 bytes

                ushort selector = ToLittleEndian(br.ReadUInt16());

                // Get number of records and offset byte value, from where font descriptions start

                ushort records = ToLittleEndian(br.ReadUInt16());
                ushort offset = ToLittleEndian(br.ReadUInt16());

                // Get the correct name record

                NameRecord nameRecord = SeekCorrectNameRecord(br, records);

                if (nameRecord != null)
                {
                    // Get the full font name for the record

                    fontName = nameRecord.ProvideFullName(br, offset);
                }
                else
                {
                    //TODO: Exception handling
                }

                br.Close();
            }

            return fontName;
        }

        /// <summary>
        /// Seeks the correct <see cref="NameRecord"/>.
        /// </summary>
        /// <param name="br">The <see cref="BinaryReader"/> to be used for reading the font metrics name table.</param>
        /// <param name="recordCount">The count of name records in the font metrics name table.</param>
        /// <returns>The <see cref="NameRecord"/> providing access to the correct full font name.</returns>
        private static NameRecord SeekCorrectNameRecord(BinaryReader br, int recordCount)
        {
            for (int i = 0; i < recordCount; i++)
            {
                NameRecord record = new NameRecord(br);

                if (record.IsWindowsUnicodeFullFontName)
                {
                    return record;
                }
            }

            return null;
        }

        /// <summary>
        /// Loads the font metrics name table as raw bytes from the font in the <paramref name="fontCollection"/>.
        /// </summary>
        /// <param name="fontCollection">The font <see cref="PrivateFontCollection"/>.</param>
        /// <returns>The metrics table as raw bytes.</returns>
        private static byte[] LoadFontMetricsNameTable(PrivateFontCollection fontCollection)
        {
            byte[] fontData = null;

            // Create dummy bitmap to generate a graphics device context

            using (Graphics g = Graphics.FromImage(new Bitmap(1, 1)))
            {
                // Get the device context

                IntPtr hdc = g.GetHdc();

                using (FontFamily family = fontCollection.Families[0])
                {
                    // Create handle to the font and load it into the device context

                    IntPtr fontHandle = new Font(family, 10f).ToHfont();
                    SelectObject(hdc, fontHandle);

                    // First determine the amount of bytes in the font metrics name table

                    uint byteCount = GetFontData(hdc, nameTableKey, 0, fontData, 0);

                    // Now init the byte array and load the data by calling GetFontData once again

                    fontData = new byte[byteCount];
                    GetFontData(hdc, nameTableKey, 0, fontData, byteCount);

                    // Release font handle

                    DeleteObject(fontHandle);
                }

                g.ReleaseHdc(hdc);
            }

            return fontData;
        }

        /// <summary>
        /// Helper function to convert any <see cref="ushort"/> value in font metrics name table to little endian,
        ///  because all stored in big endian.
        /// </summary>
        /// <param name="value">The value to be converted into little endian byte order.</param>
        /// <returns>The corresponding <see cref="ushort"/> value in big endian byte order.</returns>
        private static ushort ToLittleEndian(ushort value)
        {
            if (BitConverter.IsLittleEndian)
            {
                byte[] bytes = BitConverter.GetBytes(value);
                Array.Reverse(bytes);

                return BitConverter.ToUInt16(bytes, 0);
            }

            return value;
        }

        /// <summary>
        /// Converts the raw bytes to a <see cref="string"/> by using UTF-16BE Encoding (code page 1201).
        /// </summary>
        /// <param name="bytes">The raw bytes.</param>
        /// <returns>The converted <see cref="string"/>.</returns>
        private static string BytesToString(byte[] bytes)
        {
            // Use UTF-16BE (Unicode big endian) code page

            return Encoding.GetEncoding(1201).GetString(bytes);
        }

        #endregion

        #region TTF NameRecord Class

        /// <summary>
        /// Encapsulates a name record as specified in True Type font specification for the font metrics name table
        /// http://www.microsoft.com/typography/otspec/name.htm
        /// </summary>
        private class NameRecord
        {
            private ushort platformId;
            private ushort encodingId;
            private ushort languageId;
            private ushort nameId;

            /// <summary>
            /// Gets length of the full font name as specified in the font file.
            /// </summary>
            /// <value>
            /// The length of the full font name name.
            /// </value>
            internal ushort NameLength { get; private set; }

            /// <summary>
            /// Gets the byte offset value, where the full font name information is stored within the font metrics table .
            /// </summary>
            /// <value>
            /// The byte offset value.
            /// </value>
            internal ushort ByteOffset { get; private set; }

            /// <summary>
            /// Gets a value indicating whether this <see cref="NameRecord"/> represents a Windows Unicode full font name.
            /// </summary>
            /// <value>
            ///     <c>true</c> if this <see cref="NameRecord"/> represents a Windows Unicode full font name; otherwise, <c>false</c>.
            /// </value>
            internal bool IsWindowsUnicodeFullFontName
            {
                get
                {
                    // platformId = 3 => Windows
                    // encodingId = 1 => Unicode BMP (UCS-2)
                    // nameId = 4 => full font name

                    return platformId == 3 && encodingId == 1 && nameId == 4;
                }
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="NameRecord"/> class.
            /// </summary>
            /// <param name="br">The <see cref="BinaryReader"/> for interpretation of the bytes.</param>
            internal NameRecord(BinaryReader br)
            {
                // Read the unsigned 16-bit integers and convert to little endian

                platformId = ToLittleEndian(br.ReadUInt16());
                encodingId = ToLittleEndian(br.ReadUInt16());

                // Only read to advance reader position by 2 bytes

                languageId = ToLittleEndian(br.ReadUInt16());

                nameId = ToLittleEndian(br.ReadUInt16());
                NameLength = ToLittleEndian(br.ReadUInt16());
                ByteOffset = ToLittleEndian(br.ReadUInt16());
            }

            internal string ProvideFullName(BinaryReader br, int recordOffset)
            {
                // Search the start position of the font name

                int totalOffset = recordOffset + ByteOffset;
                br.BaseStream.Seek(totalOffset, SeekOrigin.Begin);

                // Now read the amount of bytes specified in the name record
                // and convert to a string

                byte[] nameBytes = br.ReadBytes(NameLength);
                return BytesToString(nameBytes);
            }

            /// <summary>
            /// Returns a <see cref="System.String"/> that represents this instance.
            /// </summary>
            /// <returns>
            /// A <see cref="System.String"/> that represents this instance.
            /// </returns>
            public override string ToString()
            {
                return string.Format("NameRecord - Key; Platform ID = {0}, Encoding ID = {1}, Name ID = {2}",
                             platformId.ToString(), encodingId.ToString(), nameId.ToString());
            }
        }
        #endregion
    }



}
