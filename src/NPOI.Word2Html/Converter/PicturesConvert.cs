using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace NPOI.Word2Html.Converter
{

    public class PicturesConvert
    {

        /// <summary>
        ///     图片处理
        /// </summary>
        /// <param name="myDocx"></param>
        ///// <param name="isImgUploadAliYun">图片是否上传阿里云</param>
        /// <returns></returns>
        public async Task< List<PicInfo>> PicturesHandleAsync(XWPFDocument myDocx/*, bool isImgUploadAliYun = false*/)
        {
            var picInfoList = new List<PicInfo>();
            var picturesList = myDocx.AllPictures;
            foreach (var pictures in picturesList)
            {
                var pData = pictures.Data;
                var picPackagePart = pictures.GetPackagePart();
                var picPackageRelationship = pictures.GetPackageRelationship();
                var picInfo = new PicInfo
                {
                    Id = picPackageRelationship.Id,
                    PicType = picPackagePart.ContentType
                };


                //try
                //{
                //    if (isImgUploadAliYun)
                //    {
                //        //阿里云上传图片方法
                //        var url = await _binaryObjectManager.SaveAsync(new BinaryObject
                //        { Bytes = pData, FileName = pictures.FileName, FileType = picInfo.PicType });
                //        picInfo.Url = url;
                //    }
                //}
                //catch (Exception)
                //{
                //    // ignored
                //}

                if (string.IsNullOrWhiteSpace(picInfo.Url))
                    picInfo.Url = $"data:{picInfo.PicType};base64,{Convert.ToBase64String(pData)}";
                //先把pData传阿里云得到url  如果有其他方式传改这里 或者转base64

                picInfoList.Add(picInfo);
            }

            return picInfoList;
        }

    }
}