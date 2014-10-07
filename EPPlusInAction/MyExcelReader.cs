using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using NetOffice.VBIDEApi.Enums;

namespace ExcelReportGenerator
{
    class MyExcelReader
    {
        public string Page_Load(string name)
        {
            string fName = name;

            bool bFindBOF = true;

            int col = 0;
            int row = 0;
            int totalCols = 10;

            int j = 0;
            if ((string.IsNullOrEmpty(fName)))
            {
                return String.Empty;
            }
            //Read XLS file
            try
            {
                using (var fStream = new FileStream(fName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var bReader = new BinaryReader(fStream))
                    {
                        while (j < fStream.Length - 1)
                        {
                            UInt16 uBuffer;
                            if (bFindBOF)
                            {
                                //Get BOF Record
                                while (j < fStream.Length - 3 & bFindBOF == true)
                                {
                                    uBuffer = bReader.ReadUInt16();
                                    //BOF = 0x0009
                                    if (uBuffer == 0x0009)
                                    {
                                        bFindBOF = false;
                                    }
                                    j = j + 2;
                                }
                                if (bFindBOF == false)
                                {
                                    //Get Record Length
                                    uBuffer = bReader.ReadUInt16();
                                    //Response += ("Length = " + uBuffer.ToString() + " | Offset = " + j.ToString());
                                    j = j + 2;
                                    //Get Record Body
                                   // Response += ((uBuffer > 0 ? " | Body = " : string.Empty));
//                                    for (var i = 1; i <= uBuffer; i++)
//                                    {
//                                        bBuffer = bReader.ReadByte();
//                                     //   Response += (GetByteHex(bBuffer) + " | ");
//                                    }
                                    var id = uBuffer;

                                    if (id == 0x004)
                                    {
                                        var stringArray = bReader.ReadBytes(uBuffer);

                                        var g = Encoding.Default.GetString(stringArray.Where(it => it != 0).Select(it => it).ToArray());

                                        Debug.WriteLine(g);

                                    }


                                  //  Response += ("<br/>");
                                    j = j + uBuffer;
                                }
                                else
                                {
                                    j = j + 2;
                                }
                            }
                            else
                            {
                                //Get Record ID
                                uBuffer = bReader.ReadUInt16();
                                //EOF = 0x000A
                                if (uBuffer == 10)
                                {
                                    bFindBOF = true;
                                }
                        
                                var id = uBuffer;

                                j = j + 2;
                                //Get Record Length
                                uBuffer = bReader.ReadUInt16();
                               // Response += ("Length = " + uBuffer.ToString());
                                j = j + 2;
                                //Get Record Body

                                var stringArray = bReader.ReadBytes(uBuffer);

                                if (id == 0x004)
                                {
                                //LABEL
                                //
                               
                              
                                // BIFF2:
                                //Offset Size Contents
                                //0 2 Index to row
                                //2 2 Index to column
                                //4 3 Cell attributes (➜2.5.13)
                                //7 var. Byte string, 8-bit string length (➜2.5.2)

                                    var jjj = FromTo(stringArray, 8, stringArray.Length);

                                    var g = Encoding.Default.GetString(jjj);
                                    
                                    var IndexToRow = FromTo(stringArray, 0, 2);
                                    var IndexToColumn = FromTo(stringArray, 2, 4);

                                    tttt(col, totalCols, ref row, g);

                                }
                                else if (id == 0x0024) //COLWIDTH

                                {
                                    var h = 6;
                                    ;
                                }
                                else if (id == 0x0000) //Dimensions
                                {
                                   // var Dimensions = bReader.ReadBytes(uBuffer);
                                }
                                else if (id == 0x0003) //Number
                                {
                                    var fromTo = FromTo(stringArray, 7, 15);

                                    var IndexToRow = FromTo(stringArray, 0, 2);
                                    var IndexToColumn = FromTo(stringArray, 2, 4);

                                    var number = BitConverter.ToDouble(fromTo, 0);

                                    tttt(col, totalCols, ref row, number);
                                }
                                else
                                {
                                    var k = 9;
                                }

                                //Response += ("<br/>");
                                j = j + uBuffer;
                            }
                        }

                        //  ("*ID and LENGTH take 4 bytes, " + "so total bytes per record is 4 + Length.");
                        bReader.Dispose();
                    }
                    fStream.Close();
                }
            }
            catch (Exception ex)
            {
                //Response += ("xls-check::Page_Load: exception: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }

         //   return Response;    
            return string.Empty;
        }

        private static int tttt(int col, int totalCols, ref int row, object s)
        {
            Debug.Write(s + " ");

            col++;

            if (col == totalCols)
            {
                col = 0;
                row++;

                Debug.WriteLine("");
            }
            return col;
        }

        private byte[] FromTo(byte[] y, int from, int to)
        {
            var length = to - from;

            var array = new byte[length];

            int j = 0;

            for (int i = from; i < to; i++)
            {
                array[j++] = y[i];
            }

            return array;
        }

//        private string GetUInt16Hex(UInt16 pValue)
//        {
//            string sResult = "0x";
//            double dTemp = 0;
//            //16^3
//            dTemp = Math.Floor((double) (pValue / 4096));
//            sResult = sResult + GetHexLetter(dTemp).ToString();
//            pValue = (ushort) (pValue - (dTemp * 4096));
//            //16^2
//            dTemp = Math.Floor((double) (pValue / 256));
//            sResult = sResult + GetHexLetter(dTemp).ToString();
//            pValue = (ushort) (pValue - (dTemp * 256));
//            //16^1
//            dTemp = Math.Floor((double) (pValue / 16));
//            sResult = sResult + GetHexLetter(dTemp).ToString();
//            pValue = (ushort) (pValue - (dTemp * 16));
//            //16^0
//            sResult = sResult + GetHexLetter(pValue).ToString();
//            return sResult;
//        }

//        private string GetByteHex(UInt16 pValue)
//        {
//            string sResult = "0x";
//            double dTemp = 0;
//            //16^1
//            dTemp = Math.Floor((double) (pValue / 16));
//            sResult = sResult + GetHexLetter(dTemp).ToString();
//            pValue = (ushort) (pValue - (dTemp * 16));
//            //16^0
//            sResult = sResult + GetHexLetter(pValue).ToString();
//            return sResult;
//        }

//        private string GetHexLetter(double pValue)
//        {
//            string sResult = string.Empty;
//            if (pValue < 10)
//            {
//                sResult = pValue.ToString();
//            }
//            else if (pValue == 10)
//            {
//                sResult = "A";
//            }
//            else if (pValue == 11)
//            {
//                sResult = "B";
//            }
//            else if (pValue == 12)
//            {
//                sResult = "C";
//            }
//            else if (pValue == 13)
//            {
//                sResult = "D";
//            }
//            else if (pValue == 14)
//            {
//                sResult = "E";
//            }
//            else if (pValue == 15)
//            {
//                sResult = "F";
//            }
//            return sResult;
//        }
        
    }
}