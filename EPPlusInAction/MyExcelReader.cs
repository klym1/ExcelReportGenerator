using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReportGenerator
{
    class MyExcelReader
    {
        private object[,] _result;

        public object[,] Page_Load(string name)
        {
            string fName = name;

            bool bFindBOF = true;
            
            int j = 0;
                
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
                                while (j < fStream.Length - 3 & bFindBOF)
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

                                    var id = uBuffer;

                                    if (id == 0x004)
                                    {
                                        var stringArray = bReader.ReadBytes(uBuffer);

                                        var g = Encoding.Default.GetString(stringArray.Where(it => it != 0).Select(it => it).ToArray());

                                        Debug.WriteLine(g);

                                    }
                                    
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

                                    var row = BitConverter.ToInt16(FromTo(stringArray, 0, 2), 0);
                                    var col = BitConverter.ToInt16(FromTo(stringArray, 2, 4), 0);

                                    _result[row, col] = Encoding.Default.GetString(jjj);

                                }
                                else if (id == 0x0024) //COLWIDTH
                                {
                                    //ignore
                                }
                                else if (id == 0x0000) //Dimensions
                                {
                                    //0 2 Index to first used row
                                    //2 2 Index to last used row, increased by 1
                                    //4 2 Index to first used column
                                    //6 2 Index to last used column, increased by 1

                                    var rows = BitConverter.ToInt16(FromTo(stringArray, 2, 4), 0);
                                    var cols = BitConverter.ToInt16(FromTo(stringArray, 6, 8), 0);

                                    _result = new object[rows, cols];
                                }
                                else if (id == 0x0003) //Number
                                {
                                    var fromTo = FromTo(stringArray, 7, 15);

                                    var row = BitConverter.ToInt16(FromTo(stringArray, 0, 2), 0);
                                    var col = BitConverter.ToInt16(FromTo(stringArray, 2, 4), 0);

                                    var number = BitConverter.ToDouble(fromTo, 0);

                                    _result[row, col] = number;

                                }
                                

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
            return _result;
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
    }
}