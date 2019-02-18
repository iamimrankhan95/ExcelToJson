using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToJson.Models
{
    public class ProjectData
    {
        public string PROJECTID { get; set; }
        public string PROJECTNAME { get; set; }
        public string CODE_PART1 { get; set; }
        public string CODE_PART2 { get; set; }
        public string PROJECTTITLE { get; set; }
        public string PROJECTTITLEBANGLA { get; set; }
        public string SCHEMENAME_ENG { get; set; }
        public string SCHEMENAME_BEN { get; set; }
        public string PACKAGECODE { get; set; }
        public string PACKAGEID { get; set; }
        public string SCHEMEID { get; set; }
        public string COMPONENTSUBHEADID { get; set; }
        public string COMPONENTSUBHEADNAME { get; set; }
        public string UPAZILAID { get; set; }
        public string ROADID { get; set; }
        public string ROADLENGTH { get; set; }
        public string DISTRICTID { get; set; }
        public string ISACTIVE { get; set; }
        public DateTime ACOMPLETIONDATE { get; set; }
        public string COMPONENTNAME { get; set; }
        public string CONTRACTORNAME { get; set; }
        public DateTime CONTRACTSIGNDATE { get; set; }
        public string FINANCIALYEAR { get; set; }
        public string SCHEMECODE { get; set; }
        public string PHYSICALPROGGRESS { get; set; }
        public string STATUS { get; set; }
    }
}
