using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Sitecore.DataExchange.Plugins;
using Sitecore.DataExchange.Providers.DynamicsCrm.MarketingListMemberships;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Sitecore.Support.DataExchange.Providers.DynamicsCrm.MarketingListMemberships
{
  public class ReadMarketingListMemberships: Sitecore.DataExchange.Providers.DynamicsCrm.MarketingListMemberships.ReadMarketingListMemberships
  {
    public override IEnumerable<MembershipModel> ReadListMemberships(IOrganizationService service, ReadObjectsSettings settings)
    {
      QueryExpression query = new QueryExpression
      {
        ColumnSet = new ColumnSet("listmemberid", "entityid", "listid"),
        EntityName = "listmember",
        Distinct = true
      };
      query.PageInfo = new PagingInfo
      {
        PageNumber = 1,
        Count = settings.PageSize,
        PagingCookie = null
      };
      query.LinkEntities.Add(new LinkEntity("listmember", "contact", "entityid", "contactid", JoinOperator.Inner));
      Guid id;
      while (true)
      {
        new RetrieveMultipleRequest().Query = query;
        EntityCollection results = service.RetrieveMultiple(query);
        if (results == null || !results.Entities.Any())
        {
          break;
        }
        foreach (Entity entity in results.Entities)
        {
          MembershipModel membershipModel = new MembershipModel();
          id = ((EntityReference)entity.Attributes["entityid"]).Id;
          membershipModel.EntityId = id.ToString();
          id = ((EntityReference)entity.Attributes["listid"]).Id;
          membershipModel.RelatedEntityId = id.ToString();
          yield return membershipModel;
        }
        if (!results.MoreRecords)
        {
          break;
        }
        query.PageInfo.PageNumber++;
        query.PageInfo.PagingCookie = results.PagingCookie;
      }
      List<Entity> list2 = ReadDynamicLists(service);
      List<Entity>.Enumerator enumerator2 = list2.GetEnumerator();
      try
      {
        while (enumerator2.MoveNext())
        {
          Entity list = enumerator2.Current;
          if (list.GetAttributeValue<bool>("type"))
          {
            FetchExpression query2 = new FetchExpression(list.GetAttributeValue<string>("query"));
            EntityCollection entityCollection = service.RetrieveMultiple(query2);
            foreach (Entity entity2 in entityCollection.Entities)
            {
              MembershipModel membershipModel2 = new MembershipModel();
              id = entity2.Id;
              membershipModel2.EntityId = id.ToString();
              id = list.Id;
              membershipModel2.RelatedEntityId = id.ToString();
              yield return membershipModel2;
            }
          }
        }
      }
      finally
      {
        ((IDisposable)enumerator2).Dispose();
      }
      enumerator2 = default(List<Entity>.Enumerator);
    }
  }
}