// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using Ardalis.Specification;

namespace Application.Call.Specifications
{
    public class CallGetByGraphIdSpecification : Specification<Domain.Entities.Call>
    {
        public CallGetByGraphIdSpecification(string graphCallId)
        {
            Query.Where(x => x.GraphId == graphCallId);
        }
    }
}