using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using ExcelToJson.Models;
using System.IO;
using OfficeOpenXml;
using System.Text;
using Microsoft.AspNetCore.Hosting;

namespace ExcelToJson.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        private IHostingEnvironment _env;
        public ValuesController(IHostingEnvironment env)
        {
            _env = env;
        }
        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values
        [HttpGet("excel/getJson")]
        public ActionResult GetJson()
        {
            Utility utility = new Utility();
            var webRoot = _env.WebRootPath;
            FileInfo file = new FileInfo(System.IO.Path.Combine(webRoot, "Project Data.xlsx"));
            List<ProjectData> projectDataList = new List<ProjectData>();
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {


                    for (int i = 1; i <= package.Workbook.Worksheets.Count; i++)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                        int rowCount = worksheet.Dimension.Rows;
                        int ColCount = worksheet.Dimension.Columns;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            ProjectData projectData = new ProjectData();

                            int counter = 1;

                            if (worksheet.Cells[row, 1].Value != null)
                            {
                                projectData.PROJECTID = worksheet.Cells[row, 1].Value.ToString();
                            }
                            if (worksheet.Cells[row, 2].Value != null)
                            {
                                projectData.CODE_PART1 = worksheet.Cells[row, 2].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 3].Value != null)
                            {
                                projectData.CODE_PART2 = worksheet.Cells[row, 3].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 4].Value != null)
                            {
                                projectData.PROJECTTITLE = worksheet.Cells[row, 4].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 5].Value != null)
                            {
                                projectData.PROJECTNAME = worksheet.Cells[row, 5].Value.ToString();
                            }
                            if (worksheet.Cells[row, 6].Value != null)
                            {
                                projectData.SCHEMENAME_ENG = worksheet.Cells[row, 6].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 7].Value != null)
                            {
                                projectData.SCHEMENAME_BEN = worksheet.Cells[row, 7].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 8].Value != null)
                            {
                                projectData.PACKAGEID = worksheet.Cells[row, 8].Value.ToString(); 
                            }

                            if (worksheet.Cells[row, 9].Value != null)
                            {
                                projectData.PACKAGECODE = worksheet.Cells[row, 9].Value.ToString();
                            }
                            if (worksheet.Cells[row, 10].Value != null)
                            {
                                projectData.COMPONENTSUBHEADID = worksheet.Cells[row, 10].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 11].Value != null)
                            {
                                projectData.COMPONENTSUBHEADNAME = worksheet.Cells[row, 11].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 12].Value != null)
                            {
                                projectData.UPAZILAID = worksheet.Cells[row, 12].Value.ToString();
                            }
                            if (worksheet.Cells[row, 13].Value != null)
                            {
                                projectData.ROADID = worksheet.Cells[row, 13].Value.ToString();
                            }
                            if (worksheet.Cells[row, 14].Value != null)
                            {
                                projectData.ROADLENGTH = worksheet.Cells[row, 14].Value.ToString();
                            }
                            if (worksheet.Cells[row, 15].Value != null)
                            {
                                projectData.DISTRICTID = worksheet.Cells[row, 15].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 16].Value != null)
                            {
                                projectData.ACOMPLETIONDATE = worksheet.Cells[row, 16].Value.ToString();
                            }
                            if (worksheet.Cells[row, 17].Value != null)
                            {
                                projectData.CONTRACTORNAME = worksheet.Cells[row, 17].Value.ToString();
                            }
                            if (worksheet.Cells[row, 18].Value != null)
                            {
                                projectData.CONTRACTSIGNDATE = worksheet.Cells[row, 18].Value.ToString();
                            }
                            if (worksheet.Cells[row, 19].Value != null)
                            {
                                projectData.FINANCIALYEAR = worksheet.Cells[row, 19].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 20].Value != null)
                            {
                                projectData.SCHEMECODE = worksheet.Cells[row, 20].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 21].Value != null)
                            {
                                projectData.PHYSICALPROGGRESS = worksheet.Cells[row, 21].Value.ToString();
                            }
                            if (worksheet.Cells[row, 22].Value != null)
                            {
                                projectData.STATUS = worksheet.Cells[row, 22].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 23].Value != null)
                            {
                                projectData.COMMENCEMENTDATE = worksheet.Cells[row, 23].Value.ToString();
                            }
                            if (worksheet.Cells[row, 24].Value != null)
                            {
                                projectData.REMARKS = worksheet.Cells[row, 24].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 25].Value != null)
                            {
                                projectData.FINANCIALPROGRESS = worksheet.Cells[row, 25].Value.ToString(); 
                            }
                            if (worksheet.Cells[row, 26].Value != null)
                            {
                                projectData.LATLONG = worksheet.Cells[row, 26].Value.ToString();
                            }

                            projectDataList.Add(projectData);
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                return Ok("Some error occured while importing." + ex.Message);
            }
            return Ok(projectDataList);

        }

        //https://www.talkingdotnet.com/import-export-xlsx-asp-net-core/
        //export a excel file 
        /*
                [HttpGet]
                [Route("Export")]
                public string Export()
                {
                    string sWebRootFolder = _hostingEnvironment.WebRootPath;
                    string sFileName = @"demo.xlsx";
                    string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
                    FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
                    if (file.Exists)
                    {
                        file.Delete();
                        file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
                    }
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        // add a new worksheet to the empty workbook
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employee");
                        //First add the headers
                        worksheet.Cells[1, 1].Value = "ID";
                        worksheet.Cells[1, 2].Value = "Name";
                        worksheet.Cells[1, 3].Value = "Gender";
                        worksheet.Cells[1, 4].Value = "Salary (in $)";

                        //Add values
                        worksheet.Cells["A2"].Value = 1000;
                        worksheet.Cells["B2"].Value = "Jon";
                        worksheet.Cells["C2"].Value = "M";
                        worksheet.Cells["D2"].Value = 5000;

                        worksheet.Cells["A3"].Value = 1001;
                        worksheet.Cells["B3"].Value = "Graham";
                        worksheet.Cells["C3"].Value = "M";
                        worksheet.Cells["D3"].Value = 10000;

                        worksheet.Cells["A4"].Value = 1002;
                        worksheet.Cells["B4"].Value = "Jenny";
                        worksheet.Cells["C4"].Value = "F";
                        worksheet.Cells["D4"].Value = 5000;

                        package.Save(); //Save the workbook.
                    }
                    return URL;
                }*/

        //read from a excel
        /*
                [HttpGet]
                [Route("Import")]
                public string Import()
                {
                    string sWebRootFolder = _hostingEnvironment.WebRootPath;
                    string sFileName = @"demo.xlsx";
                    FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
                    try
                    {
                        using (ExcelPackage package = new ExcelPackage(file))
                        {
                            StringBuilder sb = new StringBuilder();
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                            int rowCount = worksheet.Dimension.Rows;
                            int ColCount = worksheet.Dimension.Columns;
                            bool bHeaderRow = true;
                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= ColCount; col++)
                                {
                                    if (bHeaderRow)
                                    {
                                        sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                                    }
                                    else
                                    {
                                        sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                                    }
                                }
                                sb.Append(Environment.NewLine);
                            }
                            return sb.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        return "Some error occured while importing." + ex.Message;
                    }
                }
        */


        // GET api/values/5
        [HttpGet("{id}")]
        public ActionResult<string> Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
