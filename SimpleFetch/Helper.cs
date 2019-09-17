using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Reflection;
using System.Linq.Expressions;
using Technosoft.Yana.XRM.AI.LDMSService.Entities;
using System.Net;

namespace SimpleFetch
{
    public static class HelperCode
    {
        public static IOrganizationService GetCrmService()
        {
            string username = ConfigurationManager.AppSettings["crm:Username"].ToString();
            string password = ConfigurationManager.AppSettings["crm:Password"].ToString();
            string region = ConfigurationManager.AppSettings["crm:Region"].ToString();
            string orgUniqueId = ConfigurationManager.AppSettings["crm:OrgUniqueId"].ToString();
            string crmuri = ConfigurationManager.AppSettings["crm:Uri"].ToString();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            CrmServiceClient conn = new CrmServiceClient(username, CrmServiceClient.MakeSecureString(password), region, orgUniqueId,
                useUniqueInstance: false, useSsl: false, orgDetail: null, isOffice365: true);

            IOrganizationService service = conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

            return service;
        }
        
        public static string GetAttributeLogicalName<E, P>(this E src, Expression<Func<E, P>> expression)
        {
            var member = expression.Body as MemberExpression;
            if (member != null && member.Member is PropertyInfo)
            {
                var pi = member.Member as PropertyInfo;
                var pn = pi.Name;

                var epi = typeof(E).GetProperty(pn);
                var attr = (AttributeLogicalNameAttribute)epi.GetCustomAttribute(typeof(AttributeLogicalNameAttribute));

                if (attr == null)
                {
                    throw new System.ArgumentNullException(pn, "AttributeLogicalNameAttribute not found on property");
                }

                return attr.LogicalName;
            }

            throw new ArgumentException("Expression is not a Property");
        }

        public static EntityCollection GetEntityCollection(string FetchXML)
        {
            IOrganizationService service = HelperCode.GetCrmService();// (Username, Password);

            EntityCollection results = service.RetrieveMultiple(new FetchExpression(FetchXML));

            return results;
        }

        public static string CreateEntity(IOrganizationService service, Entity entity)
        {
            try
            {
                Guid Id = service.Create(entity);

                return Id.ToString();
            }
            catch (Exception ex)
            {

                return ex.Message.ToString();
            }
        }
        

        public static string UpdateEntity(IOrganizationService service, Entity entity)
        {
            try
            {
                service.Update(entity);

                return "Update Succeeded";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        public static IEnumerable<BusinessUnit> GetBusinessUnits(IOrganizationService service)
        {
            var bu = new BusinessUnit();

            var colsbu = new ColumnSet(
                bu.GetAttributeLogicalName(e => e.Id),
                bu.GetAttributeLogicalName(e => e.BusinessUnitId),
                bu.GetAttributeLogicalName(e => e.Name),
                bu.GetAttributeLogicalName(e => e.msdyn_companycode)
            );

            var quExp = new QueryExpression(BusinessUnit.EntityLogicalName)
            {
                ColumnSet = colsbu ?? new ColumnSet(true)
            };

            quExp.Criteria.AddCondition(bu.GetAttributeLogicalName(c => c.lxs_ismaindealer), ConditionOperator.Equal, true);

            var res = service.RetrieveMultiple(quExp).Entities;

            return res.Any() ? res.Select(e => e.ToEntity<BusinessUnit>()) : Enumerable.Empty<BusinessUnit>();
        }
        
        public static EntityCollection GetEntityCollection(IOrganizationService service, string FetchXML)
        {
            //IOrganizationService service = HelperCode.GetCrmService(Username, Password);

            EntityCollection results = service.RetrieveMultiple(new FetchExpression(FetchXML));

            return results;
        }
        public static xts_generalsetup GetGeneralSetup(IOrganizationService service, ColumnSet columnSet = null)
        {
            var query = new QueryExpression(xts_generalsetup.EntityLogicalName)
            {
                ColumnSet = columnSet ?? new ColumnSet(true),
                PageInfo = new PagingInfo { PageNumber = 1, Count = 1 },
                Criteria = { FilterOperator = LogicalOperator.And }
            };

            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            var res = service.RetrieveMultiple(query).Entities;
            if (res != null && res.Any())
            {
                return res[0].ToEntity<xts_generalsetup>();
            }

            return null;
        }
    }
}