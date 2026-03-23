using System.Collections.Generic;

namespace Microsoft.Graph
{
    public class MeetingCapability
    {
        public MeetingCapability()
        {
            ODataType = "#microsoft.graph.meetingCapability";
            AdditionalData = new Dictionary<string, object>();
        }

        public IDictionary<string, object> AdditionalData { get; set; }

        public string ODataType { get; set; }
    }
}