using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using TMS_Framework.Models;

namespace TMS_Framework.Controllers
{
    public class DocumentController : ApiController
    {
        [HttpGet]
        //Read barcode from a full path for pdf
        [Route("api/{controller}/{id}")]
        public string GetBarcode(string filePath)
        {
            // opening and getting the data from the document
            BarCode barCode = new BarCode();

            // opens the pdf and scans the barcode
            string barCodeNumber = barCode.ReadBarCode(filePath);

            // returns the barcode info
            return barCodeNumber;
        }

        [HttpGet]
        //Read barcode from a full path for pdf
        [Route("api/{controller}/{id}")]
        public string GetQRcode(string filePath)
        {
            // opening and getting the data from the document
            BarCode barCode = new BarCode();

            // opens the pdf and scans the barcode
            string qrCode = barCode.ReadQRCode(filePath);

            // returns the barcode info
            return qrCode;
        }

        [HttpPost]
        [Route("api/{controller}/{id}")]
        public IHttpActionResult Merge([FromBody] UnitOwner unitOwner)
        {
            WordAPI wordAPI = new WordAPI();

            wordAPI.ClientDataFindAndReplace(unitOwner);

            // returns the error message           
            if (!ModelState.IsValid)
            {
                HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.BadRequest);
                throw new System.Web.Http.HttpResponseException(httpResponseMessage);
            }

            return Ok();
        }

        [HttpPost]
        [Route("api/{controller}/{id}")]
        public IHttpActionResult GetTemplateMetaData([FromBody] UnitOwner unitOwner)
        {
            // opening and getting the data from the document
            WordAPI wordApi = new WordAPI();

            // the keywords inside the document populating the BaseTemplate model
            List<string> keywordList = wordApi.getExistingKeywords(unitOwner);

            // returns the filled BaseTemplate Model with the doc info in json format
            return Ok(keywordList);
        }
    }
}
