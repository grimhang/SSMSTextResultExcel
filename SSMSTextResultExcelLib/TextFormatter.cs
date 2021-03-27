using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSMSTextResultExcelLib
{
    public struct ColProperty
    {
        public int colCnt;
        public int rowCnt;
        public int[] colOriLenArr;
        public int[] colAfterLenArr;
    }

    public class TextFormatter
    {
        private int colSpaceDelimeterCntSource  = 1;        // ssms이나까 무조건 스페이스 갯수는 1개
        private int colDelimeterSpaceCntTarget;

        //public TextFormatter(int colDelimeterSpaceCntSource, int colDelimeterSpaceCntTarget)
        //{
        //    //this.colSpaceDelimeterCntSource = colDelimeterSpaceCntSource;
        //    this.colDelimeterSpaceCntTarget = colDelimeterSpaceCntTarget;            
        //}

        public TextFormatter(int colDelimeterSpaceCntTarget)
        {
            this.colDelimeterSpaceCntTarget = colDelimeterSpaceCntTarget;
        }

        public TextFormatter()
        {
        }

        #region 텍스트결과리턴
        public string OneResultFormat(string inputString)
        {
            StringBuilder oneResultFinalSB = new StringBuilder("");     // 1개 Result 가공한 후 텍스트 담을 변수

            //string oneResultSource = removeAfterNewLine(inputString.Trim());                                                                      // Result 1개의 원본테스트
            string oneResultSource = inputString.Trim().RemovePostNewLine();                                                                      // Result 1개의 원본테스트
            var oneResultArr = oneResultSource.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);          // 줄바꿈 문자 기준으로 split 하여 문자열 배열 만듬
            var oneResultLineCnt = oneResultArr.Length;                                                                       // Result 1개의 라인수

            // Result 1개의 열별 최대 길이 구하기. 
            var colProperty = calcOneResultAfterColLen(oneResultSource);

            var colCnt = colProperty.colCnt;                                    // 열 갯수 구하기
            var colOriLenArr = colProperty.colOriLenArr;                        // 열별 원본 길이 배열
            var colMaxLenArr = colProperty.colAfterLenArr;                      // 열별 최대 길이 글자수 저장할 배열

            int intStartPosition = 0;
            int intCuttingLength = 0;
            int resultSpaceCount = 0;
            string addedSpaces = "";

            for (int l = 0; l < oneResultLineCnt; l++)             //행을 루핑하여 열별 최대 길이 이용하여 최종텍스트 만들기
            {
                for (int c = 0; c < colCnt; c++)
                {
                    if (c == 0)     // 첫번째 컬럼
                    {
                        intStartPosition = 0;
                        intCuttingLength = colMaxLenArr[c];
                        resultSpaceCount = colDelimeterSpaceCntTarget;
                    } else if (c == colCnt - 1)
                    {
                        intStartPosition += colOriLenArr[c - 1] + colSpaceDelimeterCntSource;  // 2019-01-26 -1을 뺌
                        intCuttingLength = oneResultArr[l].Length - intStartPosition;
                        resultSpaceCount = 0;
                    } else
                    {
                        intStartPosition += colOriLenArr[c - 1] + colSpaceDelimeterCntSource;
                        intCuttingLength = colMaxLenArr[c];
                        resultSpaceCount = colDelimeterSpaceCntTarget;
                    }
                    addedSpaces = new string(' ', resultSpaceCount);
                    oneResultFinalSB.Append(oneResultArr[l].Substring(intStartPosition, intCuttingLength) + addedSpaces);
                }
                oneResultFinalSB.Append(Environment.NewLine);
            }
            oneResultFinalSB.Append(String.Concat(Enumerable.Repeat(Environment.NewLine, 1)));     // 1개의 레코드셋 작업하고 줄바꿈 문자를 몇개 넣어줄지
            return oneResultFinalSB.ToString(); ;
        }

        public string TotalResultFormat(string totalTextSource)
        {
            string textSourceTemp = totalTextSource.RemovePostNewLine();
            int newLineAddedCnt = 1;                                        // Result 사이의 줄바꿈 개수. 나중에 외부에서 입력 받을 예정
            StringBuilder totalResultTextFinalSB = new StringBuilder("");   // 최종 결과 Text 담을 문자열

            var totalTextSourceArr = textSourceTemp.Split(new string[] { Environment.NewLine + Environment.NewLine }, StringSplitOptions.None);
            var resultCnt = totalTextSourceArr.Length;

            for (int t = 0; t < resultCnt; t++)         // Result 1개  
            {
                var oneTextResultSource = OneResultFormat(totalTextSourceArr[t].Trim());
                totalResultTextFinalSB.Append(oneTextResultSource + String.Concat(Enumerable.Repeat(Environment.NewLine, newLineAddedCnt)));     // 1개의 레코드셋 작업하고 줄바꿈 문자를 몇개 넣어줄지
            }

            return totalResultTextFinalSB.ToString();
        }
        #endregion

        #region Datatable결과리턴
        public DataTable OneResultFormatDT(string inputString)
        {
            //string oneResultSource = removeAfterNewLine(inputString.TrimStart());
            string oneResultSource = inputString.TrimStart().RemovePostNewLine();

            var colProperty = calcOneResultAfterColLen(oneResultSource);

            var colCnt = colProperty.colCnt;                              // 열 갯수 구하기
            var colOriLenArr = colProperty.colOriLenArr;                  // 열별 원본 길이 배열
            var colMaxLenArr = colProperty.colAfterLenArr;                // 열별 최대 길이 글자수 저장할 배열

            var oneResultArr = oneResultSource.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);      // 줄바꿈 문자 기준으로 split 하여 문자열 배열 만듬
            var oneResultLineCnt = oneResultArr.Length;                                                                   // Result 1개의 라인수

            //Debug.Print($"   TotalResultFormatDS:result:{result}");
            // Result 1개의 열별 최대 길이 구하기.


            var dt = new DataTable();

            int intStartPosition = 0;
            int intCuttingLength = 0;
            int resultSpaceCount = 0;

            for (int line = 0; line < oneResultLineCnt; line++)
            {
                DataRow dr1 = dt.NewRow();

                for (int col = 0; col < colCnt; col++)
                {
                    if (col == 0)                                       // 첫번째 컬럼
                    {
                        intStartPosition = 0;
                        intCuttingLength = colMaxLenArr[col];
                        resultSpaceCount = colDelimeterSpaceCntTarget;
                    } else if (col == colCnt - 1)                       // 마지막 컬럼
                    {
                        intStartPosition += colOriLenArr[col - 1] + colSpaceDelimeterCntSource;
                        intCuttingLength = oneResultArr[line].Length - intStartPosition;
                        resultSpaceCount = 0;
                    } else
                    {
                        intStartPosition += colOriLenArr[col - 1] + colSpaceDelimeterCntSource;
                        intCuttingLength = colMaxLenArr[col];
                        resultSpaceCount = colDelimeterSpaceCntTarget;
                    }                    
                    
                    //---------------------------------------------------------------------------------------------
                    if (line == 0)     // 첫번째 줄은 컬럼명이니까 컬럼추가에 이용
                    {
                        string colName = oneResultArr[line].Substring(intStartPosition, intCuttingLength).Trim();

                        DataColumn col1 = new DataColumn();
                        col1.DataType = System.Type.GetType("System.String");
                        col1.ColumnName = colName;// "ID";
                        dt.Columns.Add(col1);
                    } 

                    if (line >= 2)
                    {
                        dr1[col] = oneResultArr[line].Substring(intStartPosition, intCuttingLength).TrimEnd();
                    }
                }

                if (line >= 2)
                {
                    dt.Rows.Add(dr1);
                }
            }
            return dt;
        }

        public DataSet TotalResultFormatDS(string totalResultTextSource, bool titleInclude = false)
        {
            //http://www.gisdeveloper.co.kr/?p=1070

            DataSet ds = new DataSet();

            //int newLineAddedCnt = 1;                                        // Result 사이의 줄바꿈 개수. 나중에 외부에서 입력 받을 예정
            //StringBuilder totalResultTextFinalSB = new StringBuilder("");   // 최종 결과 Text 담을 문자열

            //var totalTextResultSourceArr = totalResultTextSource.Trim().Split(new string[] { Environment.NewLine + Environment.NewLine }, StringSplitOptions.None);
            //var totalTextResultSourceArr = removeAfterNewLine(totalResultTextSource.TrimStart()).Split(new string[] { Environment.NewLine + Environment.NewLine }, StringSplitOptions.None);
            var totalTextResultSourceArr = totalResultTextSource.TrimStart().RemovePostNewLine().Split(new string[] { Environment.NewLine + Environment.NewLine }, StringSplitOptions.None);
            var resultCnt = totalTextResultSourceArr.Length;

            for (int result = 0; result < resultCnt; result++)         // Result 1개  
            {
                //var oneTextResultSource = OneResultFormat(totalResultTextSourceArr[t].Trim());
                //totalResultTextFinalSB.Append(oneTextResultSource + String.Concat(Enumerable.Repeat(Environment.NewLine, newLineAddedCnt)));     // 1개의 레코드셋 작업하고 줄바꿈 문자를 몇개 넣어줄지
                //string trimmedOneTextResultSource = totalTextResultSourceArr[result].Trim();
                //string trimmedOneTextResultSource = removeAfterNewLine(totalTextResultSourceArr[result].TrimStart());
                string trimmedOneTextResultSource = totalTextResultSourceArr[result].TrimStart().RemovePostNewLine();

                //string[] temp = new string[] { };
                string title = "";

                if (titleInclude)
                {
                    var temp = trimmedOneTextResultSource.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    title = temp[0].Replace("--##  ", "");

                    trimmedOneTextResultSource = trimmedOneTextResultSource.Replace(temp[0], "").TrimStart();
                }

                Debug.Print($"TotalResultFormatDS:result:{result}");

                var oneTextResultSourceDT = OneResultFormatDT(trimmedOneTextResultSource);

                if (titleInclude)
                {
                    oneTextResultSourceDT.TableName = title;
                }

                ds.Tables.Add(oneTextResultSourceDT);
            }

            //return totalResultTextFinalSB.ToString();
            return ds;
        }
        #endregion

        private ColProperty calcOneResultAfterColLen(string inputString)
        {
            try
            {
                string oneResultSource = inputString.TrimStart().RemovePostNewLine();

                var inputStringArr = oneResultSource.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                var lineCnt = inputStringArr.Length;

                string strDashLine = inputStringArr[1];        // 2번째 줄을 기준으로 컬럼의 길이 계산하기 위해
                var strColArr = strDashLine.Split(
                        new string[] { new string(' ', colSpaceDelimeterCntSource) }, StringSplitOptions.None
                    );        // 스페이스 갯수로 split해서 열갯수 세기

                int colCnt = strColArr.Length;  // 컬럼 갯수

                int[] colLenArr = new int[colCnt];

                // Length Per column
                for (int i = 0; i < colCnt; i++)
                {
                    colLenArr[i] = strColArr[i].Length;         // 열의 원래 길이
                }

                int intStartPosition = 0;
                int intCuttingLength = 0;

                int[] colAfterLenArr = new int[colCnt];

                for (int line = 0; line < lineCnt; line++)    //행을 루핑
                {
                    if (line != 1)                                                     // 2번째 행은 원본 컬럼의 Max길이 이기 때문에 최대길이 구하는 로직에서 제외
                    {
                        for (int col = 0; col < colCnt; col++)        //열을 루핑하여 컬럼당 최대 길이 구하기
                        {
                            if (col == 0)                                             // 첫번째 컬럼일 경우
                            {
                                intStartPosition = 0;
                                intCuttingLength = colCnt == 1 ? inputStringArr[line].TrimEnd().Length : colLenArr[col];      // 컬럼수가 1개일때는 라인의 길이가 컬럼의 길이
                            } else if (col == colCnt - 1)                             // 마지막 컬럼일 경우
                            {
                                intStartPosition += colLenArr[col - 1] + 1;
                                intCuttingLength = inputStringArr[line].Length - intStartPosition;
                            } else
                            {
                                intStartPosition += colLenArr[col - 1] + 1;
                                intCuttingLength = colLenArr[col];
                            }

                            int intNowColLength = inputStringArr[line].Substring(intStartPosition, intCuttingLength).TrimEnd().Length;

                            if (intNowColLength > colAfterLenArr[col])
                            {
                                colAfterLenArr[col] = intNowColLength;
                            }
                        }
                    }

                    Debug.Print($"  calcOneResultAfterColLen:Row:{line}");
                }

                return (new ColProperty()
                {
                    colCnt = colCnt
    ,
                    rowCnt = lineCnt
    ,
                    colOriLenArr = colLenArr
    ,
                    colAfterLenArr = colAfterLenArr
                });
            } catch (IndexOutOfRangeException err)
            {
                //System.IndexOutOfRangeException: '인덱스가 배열 범위를 벗어났습니다.'
                throw new IndexOutOfRangeException($"데이터가 잘못되었습니다:({System.Reflection.MethodBase.GetCurrentMethod().Name})" , err);
            } 
            //catch (ArgumentOutOfRangeException err)
            //{
            //    throw new Exception($"Input data is not valid", err);
            //}

            //return (new ColProperty());
        }

        public bool TryCheckTextTotal (string inputText, out string errMsg)
        {
            // 01. 
            errMsg = null;
            return true;
        }

        public bool TryCheckTextPartial(string inputText, out string errMsg)
        {
            bool finalBool = true;

            // 01. 두번째 줄이 대시라인인지체크
            string textSource = inputText.TrimStart().RemovePostNewLine();                                                                      // Result 1개의 원본테스트
            var oneResultArr = textSource.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);          // 줄바꿈 문자 기준으로 split 하여 문자열 배열 만듬

            int dashRowIndex = textSource.StartsWith("--##  ") ? 2 : 1;

            if (!string.IsNullOrEmpty(oneResultArr[dashRowIndex].Replace("-", "").Replace(" ", "")))
            {
                finalBool = false;
                errMsg = "DashLine not exist.";
            }else
                errMsg = null;

            return finalBool;
        }
    }
}
