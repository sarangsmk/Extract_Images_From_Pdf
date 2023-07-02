/// <summary>
/// Description: Utility project/file which extracts images from a  given PDF file.
/// Created by: Sarang MK
/// Namespace: Capture_Content_From_Pdf
/// Date: July 2, 2023
/// Github: https://github.com/sarangsmk
/// </summary>

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Capture_Content_From_Pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            string pdfFilePath = @"D:\Projects\Capture_Content_From_PDF\SamplePDF.pdf"; // Specify the path to the PDF file

            try
            {
                SaveImages(pdfFilePath, @"D:\Projects\Capture_Content_From_PDF\OutputImage");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            Console.ReadLine();
        }


        /// <summary>
        /// Extracts images from a pdf, and saves them to a file.
        /// </summary>
        /// <param name="pathToPdf">Pdf file.</param>
        /// <param name="outputPath">Output path where the image will be stored.</param>
        private static void SaveImages(string pathToPdf, string outputPath)
        {
            try
            {
                string name = System.IO.Path.GetFileNameWithoutExtension(pathToPdf);
                if (!Directory.Exists(outputPath)) Directory.CreateDirectory(outputPath);

                // Get a List of Image
                List<Image> ListImage = ExtractImages(pathToPdf);

                for (int i = 0; i < ListImage.Count; i++)
                {
                    try
                    {
                        string currentName = name + i + ".jpg";

                        Bitmap bmpImage = new Bitmap(ListImage[i]);
                        bmpImage.Save(System.IO.Path.Combine(outputPath, currentName));
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Un-expected Exception - " + e.Message);
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        // <summary>
        /// Extract all images from a pdf, and store them in a list of Images.
        /// </summary>
        /// <param name="PDFSourcePath">Specify PDF Source Path</param>
        /// <returns>List</returns>
        private static List<Image> ExtractImages(string PDFSourcePath)
        {
            List<Image> ImgList = new List<Image>();

            try
            {
                RandomAccessFileOrArray RAFObj = null;
                PdfReader PDFReaderObj = null;
                PdfObject PDFObj = null;
                PdfStream PDFStremObj = null;

                RAFObj = new RandomAccessFileOrArray(PDFSourcePath);
                PDFReaderObj = new PdfReader(RAFObj, null);

                for (int i = 0; i <= PDFReaderObj.XrefSize - 1; i++)
                {
                    PDFObj = PDFReaderObj.GetPdfObject(i);

                    if ((PDFObj != null) && PDFObj.IsStream())
                    {
                        PDFStremObj = (PdfStream)PDFObj;
                        PdfObject subtype = PDFStremObj.Get(PdfName.SUBTYPE);

                        if ((subtype != null) && subtype.ToString() == PdfName.IMAGE.ToString())
                        {
                            try
                            {
                                PdfImageObject PdfImageObj =
                                    new PdfImageObject((PRStream)PDFStremObj);

                                Image ImgPDF = PdfImageObj.GetDrawingImage();

                                ImgList.Add(ImgPDF);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Un-expected Exception - " + e.Message);
                            }
                        }
                    }
                }
                PDFReaderObj.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return ImgList;
        }
    }

}
