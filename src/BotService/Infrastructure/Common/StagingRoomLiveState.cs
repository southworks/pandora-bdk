using System.Collections.Generic;

namespace Microsoft.Graph
{
    public class StagingRoomLiveState
    {
        public StagingRoomLiveState()
        {
            ODataType = "#microsoft.graph.stagingRoomLiveState";
            AdditionalData = new Dictionary<string, object>();
        }

        public IDictionary<string, object> AdditionalData { get; set; }

        public string ODataType { get; set; }
    }
}