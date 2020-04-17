using NPOI.Word2Html.Converter;
using NPOI.XWPF.UserModel;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Word2Html
{
    public class NpoiDoc
    {

        /// <summary>
        ///     Npoi处理Docx
        /// </summary>
        /// <param name="stream"></param>
        ///// <param name="isImgUploadAliYun"></param>
        /// <returns></returns>
        public async Task<string> NpoiDocx(Stream stream/*, bool isImgUploadAliYun = false*/)
        {
            var myDocx = new XWPFDocument(stream); //打开07（.docx）以上的版本的文档

            PicturesConvert picturesConvert = new PicturesConvert();
            ParaGraphConvert paraGraphConvert = new ParaGraphConvert();
            TableConvert tableConvert = new TableConvert();
            var picInfoList = await picturesConvert.PicturesHandleAsync(myDocx/*, isImgUploadAliYun*/);

            var sb = new StringBuilder();

            foreach (var para in myDocx.BodyElements)
                switch (para.ElementType)
                {
                    case BodyElementType.PARAGRAPH:
                        {
                            var paragraph = (XWPFParagraph)para;
                            sb.Append(paraGraphConvert.ParaGraphHandle(paragraph, picInfoList));

                            break;
                        }

                    case BodyElementType.TABLE:
                        var paraTable = (XWPFTable)para;
                        sb.Append(tableConvert.TableHandle(paraTable, picInfoList));
                        break;
                }


            return sb.Replace(" style=''", "").ToString();
        }
    }
}