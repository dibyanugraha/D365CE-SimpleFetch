using Microsoft.Xrm.Sdk;
using System;
using Technosoft.Yana.XRM.AI.LDMSService.Entities;

namespace SimpleFetch
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Test");
            HelperCode.GetCrmService();
            GetSomething();
            Console.ReadLine();
        }

        static void GetSomething()
        {
            var fetchXML = string.Format(
            @"
            <fetch>
                <entity name='xts_newvehicledeliveryorder' >
                <attribute name='xjp_billtocustomerid' alias='CustId' />
                <attribute name='xts_newvehicledeliveryordernumber' alias='NVDONo' />
                <link-entity name='account' from='accountid' to='xjp_billtocustomerid' link-type='inner' >
                    <attribute name='accountnumber' alias='CustNo' />
                    <link-entity name='xts_customerclass' from='xts_customerclassid' to='xts_customerclassid' link-type='inner' >
                    <attribute name='xts_accountreceiveabledim5from' alias='Dim5From' />
                    <attribute name='xts_dimension5id' alias='Dim5Id' />
                    <attribute name='xts_customerclass' alias='CustClass' />
                    </link-entity>
                </link-entity>
                </entity>
            </fetch>");

            EntityCollection results = HelperCode.GetEntityCollection(fetchXML);

            if (results.Entities.Count > 0)
            {
                foreach (var entity in results.Entities)
                {
                    if (entity.Attributes.Contains("Dim5From"))
                    {
                        Console.WriteLine(((AliasedValue)entity["CustNo"]).Value);
                        Console.WriteLine(entity.FormattedValues["CustId"]);
                        Console.WriteLine(entity.FormattedValues["Dim5Id"]);
                        Console.WriteLine(entity.FormattedValues["Dim5From"]);
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}
