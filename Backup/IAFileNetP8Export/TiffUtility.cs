using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;



namespace CNO.BPA.IAFileNetP8Export
{
   class TiffUtility
   {

      public TiffUtility()
      {
      }

      private ImageCodecInfo GetEncoderInfo(string mimeType)
      {
         ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
         for (int j = 0; j < encoders.Length; j++)
         {
            if (encoders[j].MimeType == mimeType)
               return encoders[j];
         }

         throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
      }
      public Stream JoinImages(ref List<Stream> images, ref List<Stream> replicateimages)
      {
         try
         {
            EncoderParameters ep = new EncoderParameters(2);
            ep.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
            ep.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

            Image pages = null;
            MemoryStream multiStream = new MemoryStream();
            ImageCodecInfo info = GetEncoderInfo("image/tiff");

            foreach (Stream singleImage in replicateimages)
            {
               AddImage(ref multiStream, streamToImage(singleImage), ref pages, ref ep, ref info );
            }
            foreach (Stream singleImage in images)
            {
               AddImage(ref multiStream, streamToImage(singleImage), ref pages, ref ep, ref info);
            }

            //flush and close.
            ep.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
            pages.SaveAdd(ep);
            pages.Dispose();
            ep.Dispose();


            return multiStream;

         }
         catch (Exception ex)
         {
            throw new Exception("TiffUtility.JoinImages: " + ex.Message);
         }

      }
      public void AddImage(ref MemoryStream multiStream, Image singleImage, ref Image pages, ref EncoderParameters ep, ref ImageCodecInfo info)
      {
         Guid objGuid = singleImage.FrameDimensionsList[0];
         FrameDimension objDimension = new FrameDimension(objGuid);

         //Gets the total number of frames in the .tiff file
         int PageNumber = singleImage.GetFrameCount(objDimension);

         for (int i = 0; i < PageNumber; i++)
         {
            singleImage.SelectActiveFrame(objDimension, i);

            //save the intermediate frames
            if (multiStream.Length == 0)
            {
               pages = singleImage;
               try
               {
                  pages.Save(multiStream, info, ep);
               }
               catch
               {
                  pages = ConvertToBitonal(singleImage);
                  pages.Save(multiStream, info, ep);
               }

            }
            else
            {
               ep.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
               try
               {
                  pages.SaveAdd(singleImage, ep);
               }
               catch
               {
                  using (Image bitonalImage = ConvertToBitonal(singleImage))
                  {
                     pages.SaveAdd(bitonalImage, ep);
                  }
               }
            }
         }

      }
      public static Image ConvertToBitonal(Image original)
      {
         try
         {

            Bitmap source = null;

            // If original bitmap is not already in 32 BPP, ARGB format, then convert
            if (original.PixelFormat != PixelFormat.Format32bppArgb)
            {
               source = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
               source.SetResolution(original.HorizontalResolution, original.VerticalResolution);
               using (Graphics g = Graphics.FromImage(source))
               {
                  g.DrawImageUnscaled(original, 0, 0);
               }
            }
            else
            {
               source = new Bitmap(original);
            }

            // Lock source bitmap in memory
            BitmapData sourceData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            // Copy image data to binary array
            int imageSize = sourceData.Stride * sourceData.Height;
            byte[] sourceBuffer = new byte[imageSize];
            Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, imageSize);

            // Unlock source bitmap
            source.UnlockBits(sourceData);

            // Create destination bitmap
            Bitmap destination = new Bitmap(source.Width, source.Height, PixelFormat.Format1bppIndexed);

            // Lock destination bitmap in memory
            BitmapData destinationData = destination.LockBits(new Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);

            // Create destination buffer
            imageSize = destinationData.Stride * destinationData.Height;
            byte[] destinationBuffer = new byte[imageSize];

            int sourceIndex = 0;
            int destinationIndex = 0;
            int pixelTotal = 0;
            byte destinationValue = 0;
            int pixelValue = 128;
            int height = source.Height;
            int width = source.Width;
            int threshold = 500;

            // Iterate lines
            for (int y = 0; y < height; y++)
            {
               sourceIndex = y * sourceData.Stride;
               destinationIndex = y * destinationData.Stride;
               destinationValue = 0;
               pixelValue = 128;

               // Iterate pixels
               for (int x = 0; x < width; x++)
               {
                  // Compute pixel brightness (i.e. total of Red, Green, and Blue values)
                  pixelTotal = sourceBuffer[sourceIndex + 1] + sourceBuffer[sourceIndex + 2] + sourceBuffer[sourceIndex + 3];
                  if (pixelTotal > threshold)
                  {
                     destinationValue += (byte)pixelValue;
                  }
                  if (pixelValue == 1)
                  {
                     destinationBuffer[destinationIndex] = destinationValue;
                     destinationIndex++;
                     destinationValue = 0;
                     pixelValue = 128;
                  }
                  else
                  {
                     pixelValue >>= 1;
                  }
                  sourceIndex += 4;
               }
               if (pixelValue != 128)
               {
                  destinationBuffer[destinationIndex] = destinationValue;
               }
            }

            // Copy binary image data to destination bitmap
            Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize);

            // Unlock destination bitmap
            destination.UnlockBits(destinationData);
            source.Dispose();
            // Return
            return destination;

         }
         catch (Exception ex)
         {
            throw new Exception("TiffUtility.ConvertToBitonal: " + ex.Message);
         }


      }


      public Image streamToImage(Stream streamIn)
      {
         try
         {

           return Image.FromStream(streamIn);

         }
         catch (Exception ex)
         {
            throw new Exception("TiffUtility.byteArrayToImage: " + ex.Message);
         }

      }


   }
}
