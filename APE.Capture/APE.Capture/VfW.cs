using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using NM = APE.Native.NativeMethods;

namespace APE.Capture
{
    internal class VfW
    {
        // AVI file
        private static IntPtr m_File;
        // Video stream
        private static IntPtr m_Stream;
        // Compressed stream
        private static IntPtr m_StreamCompressed;
        // Buffer to store the flipped frame before writing it to the file
        private static IntPtr m_FlippedFrameBuffer = IntPtr.Zero;
        // Width
        private static int m_Width;
        // Height
        private static int m_Height;
        // Length of a scanline
        private static int m_Stride;
        // Quality
        private static int m_Quality = -1;
        // How often to insert a key frame
        private static int m_KeyFrameEvery = 0;
        // Frame rate
        private static int m_Rate = 25;
        // Current frame in the file
        private static int m_CurrentFrame;
        // Codec used for video compression
        private static string m_Codec = "DIB ";

        public static int FrameRate
        {
            get
            {
                return m_Rate;
            }
            set
            {
                m_Rate = value;
            }
        }

        public static string Codec
        {
            get
            {
                return m_Codec;
            }
            set
            {
                m_Codec = value;
            }
        }

        public static int Quality
        {
            get
            {
                return m_Quality;
            }
            set
            {
                m_Quality = value;
            }
        }

        public static int KeyFrameEvery
        {
            get
            {
                return m_KeyFrameEvery;
            }
            set
            {
                m_KeyFrameEvery = value;
            }
        }

        static VfW()
        {
            NM.AVIFileInit();
        }

        private static int mmioFOURCC(string str)
        {
            return (
                ((int)(byte)(str[0])) |
                ((int)(byte)(str[1]) << 8) |
                ((int)(byte)(str[2]) << 16) |
                ((int)(byte)(str[3]) << 24));
        }

        public static Bitmap Open(string fileName, int width, int height, PixelFormat format)
        {
            // Make sure any previous file is closed
            Close();

            bool success = false;
            Bitmap theBitmap = new Bitmap(width, height, format);

            // Get the actual stride
            BitmapData bmpData = theBitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadOnly, theBitmap.PixelFormat);
            m_Stride = bmpData.Stride;
            theBitmap.UnlockBits(bmpData);

            try
            {
                // create new file
                if (NM.AVIFileOpen(out m_File, fileName, NM.OpenFileMode.Create | NM.OpenFileMode.Write, IntPtr.Zero) != 0)
                {
                    throw new Exception("Failed to create the file");
                }

                m_Width = width;
                m_Height = height;

                // describe new stream
                NM.AVISTREAMINFO info = new NM.AVISTREAMINFO();

                info.fccType = mmioFOURCC("vids");
                info.fccHandler = mmioFOURCC(m_Codec);
                info.dwScale = 1;
                info.dwRate = FrameRate;
                info.dwSuggestedBufferSize = m_Stride * height;

                // create stream
                if (NM.AVIFileCreateStream(m_File, out m_Stream, ref info) != 0)
                {
                    throw new Exception("Failed to create the stream");
                }

                // describe compression options
                NM.AVICOMPRESSOPTIONS options = new NM.AVICOMPRESSOPTIONS();

                options.fccHandler = mmioFOURCC(m_Codec);
                options.dwQuality = Quality * 100;
                options.dwKeyFrameEvery = KeyFrameEvery;
                if (options.dwKeyFrameEvery == 0)
                {
                    options.dwFlags = NM.AviCompression.AVICOMPRESSF_VALID;
                }
                else
                {
                    options.dwFlags = NM.AviCompression.AVICOMPRESSF_KEYFRAMES | NM.AviCompression.AVICOMPRESSF_VALID;
                }

                // Create the compressed stream
                if (NM.AVIMakeCompressedStream(out m_StreamCompressed, m_Stream, ref options, IntPtr.Zero) != 0)
                {
                    throw new Exception("Failed  to create the compressed stream");
                }

                // Create the header for the frame format
                NM.BITMAPINFOHEADER bitmapInfoHeader = new NM.BITMAPINFOHEADER();

                bitmapInfoHeader.biSize = Marshal.SizeOf(typeof(NM.BITMAPINFOHEADER));
                bitmapInfoHeader.biWidth = width;
                bitmapInfoHeader.biHeight = height;
                bitmapInfoHeader.biPlanes = 1;
                bitmapInfoHeader.biBitCount = (short)Image.GetPixelFormatSize(format);
                bitmapInfoHeader.biSizeImage = 0;
                bitmapInfoHeader.biCompression = 0; // BI_RGB

                // Write the header
                if (NM.AVIStreamSetFormat(m_StreamCompressed, 0, ref bitmapInfoHeader, bitmapInfoHeader.biSize) != 0)
                {
                    throw new Exception("Failed to set the compressed stream format");
                }

                // alloc unmanaged memory for the flipped frame buffer
                m_FlippedFrameBuffer = Marshal.AllocHGlobal(m_Stride * height);

                if (m_FlippedFrameBuffer == IntPtr.Zero)
                {
                    throw new Exception("Failed to allocate memory for the flipped frame buffer");
                }

                m_CurrentFrame = 0;
                success = true;
            }
            finally
            {
                if (!success)
                {
                    Close();
                }
            }

            return theBitmap;
        }

        public static void Close()
        {
            // Free unmanaged memory
            if (m_FlippedFrameBuffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(m_FlippedFrameBuffer);
                m_FlippedFrameBuffer = IntPtr.Zero;
            }

            // Release the compressed stream
            if (m_StreamCompressed != IntPtr.Zero)
            {
                NM.AVIStreamRelease(m_StreamCompressed);
                m_StreamCompressed = IntPtr.Zero;
            }

            // Release the stream
            if (m_Stream != IntPtr.Zero)
            {
                NM.AVIStreamRelease(m_Stream);
                m_Stream = IntPtr.Zero;
            }

            // Release the file
            if (m_File != IntPtr.Zero)
            {
                NM.AVIFileRelease(m_File);
                m_File = IntPtr.Zero;
            }
        }

        public static void AddFrame(Bitmap frameImage)
        {
            // Lock the bitmap
            BitmapData imageData = frameImage.LockBits(new Rectangle(0, 0, m_Width, m_Height), ImageLockMode.ReadOnly, frameImage.PixelFormat);

            // Copy and flip the frame
            IntPtr pSourceScanline = imageData.Scan0 + m_Stride * (m_Height - 1);
            IntPtr pDestinationScanline = m_FlippedFrameBuffer;
            IntPtr pStride = new IntPtr(m_Stride);

            for (int y = 0; y < m_Height; y++)
            {
                NM.memcpy(pDestinationScanline, pSourceScanline, pStride);
                pDestinationScanline += m_Stride;
                pSourceScanline -= m_Stride;
            }

            // Unlock the bitmap
            frameImage.UnlockBits(imageData); ;

            // Write the frame to the stream
            if (NM.AVIStreamWrite(m_StreamCompressed, m_CurrentFrame, 1, m_FlippedFrameBuffer, m_Stride * m_Height, 0, IntPtr.Zero, IntPtr.Zero) != 0)
            {
                throw new Exception("Failed to add frame");
            }

            m_CurrentFrame++;
        }
    }
}
