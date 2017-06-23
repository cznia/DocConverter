using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Swashbuckle.Swagger.Annotations;
using Aspose.Words;
using Aspose.Cells;
using Aspose.Slides;
using System.Threading.Tasks;

namespace DocConverter.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values/5
        [SwaggerOperation("GetById")]
        [SwaggerResponse(HttpStatusCode.OK)]
        [SwaggerResponse(HttpStatusCode.NotFound)]
        public async Task<string> GetAsync(int id)
        {

            string basePath = @"C:\test\";
            string baseExportPath = @"C:\test\Export\";

            string fileName = @"demo";

            switch (id)
            {
                case 1:

                    List<bool> loop = new List<bool>();

                    for (int x = 0; x <= 200; x++)
                    {
                        loop.Add(await test(x));
                    }
                    break;
                case 2:
                    Workbook xls = new Workbook(basePath + fileName + ".xlsx");
                    xls.Save(baseExportPath + "xls.pdf", Aspose.Cells.SaveFormat.Pdf);
                    xls.Save(baseExportPath + "xls.ods", Aspose.Cells.SaveFormat.ODS);
                    break;
                case 3:
                    Presentation ppt = new Presentation(basePath + fileName + ".pptx");
                    ppt.Save(baseExportPath + "ppt.pdf", Aspose.Slides.Export.SaveFormat.Pdf);
                    ppt.Save(baseExportPath + "ppt.odp", Aspose.Slides.Export.SaveFormat.Odp);
                    break;
            }



            return "value";
        }

        /// <summary>
        /// 多文件測試
        /// </summary>
        /// <param name="x"></param>
        /// <returns></returns>
        private async Task<bool> test(int x)
        {
            try
            {
                string basePath = @"C:\test\";
                string baseExportPath = @"C:\test\Export\";

                string fileName = @"demo";

                Document doc = new Document(basePath + fileName + ".docx");
                doc.Save(baseExportPath + $"doc-{x}.pdf", Aspose.Words.SaveFormat.Pdf);
                doc.Save(baseExportPath + $"doc-{x}.odt", Aspose.Words.SaveFormat.Odt);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }



        }
       
    }
}
