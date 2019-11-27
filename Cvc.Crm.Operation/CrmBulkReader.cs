using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Cvv.Crm.Operation
{
    public class CrmBulkReader
    {
        private readonly QueryExpression query;
        private readonly IOrganizationService service;
        private PagingInfo pagingInfo;
        public bool MoreRecords { get; private set; } = true;
        public int CurrentPage => pagingInfo.PageNumber;
        public int PageNumber { get; private set; }
        public int TotalCount { get; }

        public CrmBulkReader(IOrganizationService service, QueryExpression query, int count = 5000)
        {
            this.query = query;
            this.service = service;
            InitPageInfo(count);
        }

        private void InitPageInfo(int count)
        {
            pagingInfo = new PagingInfo
            {
                Count = count,
                ReturnTotalRecordCount = true,
                PageNumber = 1
            };
            query.PageInfo = pagingInfo;
        }

        public List<T> Read<T>() where T : Entity
        {
            if (!MoreRecords)
            {
                return null;
            }
            var dataCollection = service.RetrieveMultiple(query);
            pagingInfo.PageNumber++;
            pagingInfo.PagingCookie = dataCollection.PagingCookie;
            MoreRecords = dataCollection.MoreRecords;

            return dataCollection.Entities.Cast<T>().ToList();
        }

        public List<Entity> Read()
        {
            return Read<Entity>();
        }

        public List<Entity> ReadAllParallel(Func<IOrganizationService> service)
        {
            //InitPageInfo(pagingInfo.Count);
            ConcurrentBag<Entity> resultCollection = new ConcurrentBag<Entity>();
            Parallel.ForEach(Enumerable.Range(1, 4), new ParallelOptions { MaxDegreeOfParallelism = 4 }, (index) =>
             {
                 var q = new QueryExpression();
                 pagingInfo = new PagingInfo
                 {
                     Count = 5000,
                     //ReturnTotalRecordCount = true,
                     PageNumber = index
                 };
                 q.EntityName = query.EntityName;
                 q.PageInfo = pagingInfo;
                 q.Criteria = query.Criteria;
                 q.ColumnSet = query.ColumnSet;

                 var dataCollection = service().RetrieveMultiple(q);
                 foreach (var item in dataCollection.Entities)
                 {
                     resultCollection.Add(item);
                 }
             });

            return resultCollection.ToList();
        }

        public List<T> ReadAll<T>() where T : Entity
        {
            InitPageInfo(pagingInfo.Count);
            var list = new List<T>();

            while (MoreRecords)
            {
                var dataCollection = service.RetrieveMultiple(query);
                pagingInfo.PageNumber++;
                pagingInfo.PagingCookie = dataCollection.PagingCookie;
                MoreRecords = dataCollection.MoreRecords;

                list.AddRange(dataCollection.Entities.Cast<T>());
            }

            return list;
        }

        public List<Entity> ReadAll()
        {
            return ReadAll<Entity>();
        }

        public Task<List<T>> ReadAllAsync<T>() where T : Entity
        {
            return Task.Run(() => ReadAll<T>());
        }
    }
}