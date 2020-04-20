using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using NNIT.MicrosoftPlanner.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;
using NNIT.MircrosoftPlanner.Activities;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using UiPath.Shared.Activities.HTTP;

namespace NNIT.MicrosoftPlanner.Activities.Plan
{
    [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetPlannerGroup_Description))]
    public class GetPlannerGroup : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.Timeout_DisplayName))]
        [LocalizedDescription(nameof(Resources.Timeout_Description))]
        public InArgument<int> TimeoutMS { get; set; } = 60000;

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_GroupId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_GroupId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> GroupId { get; set; }


        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_CreatedDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_CreatedDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> CreatedDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_Description_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_Description_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Description { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_Mail_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_Mail_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Mail { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_Visibillity_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_Visibillity_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Visibillity { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_GroupDictionary_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_GroupDictionary_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, object>> GroupDictionary { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlannerGroup_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlannerGroup_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetPlannerGroup()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetPlannerGroup, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }


        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (GroupId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(GroupId)));
           
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            string authToken = objectContainer.Get<string>();
            int timeout = TimeoutMS.Get(context);
            string groupId = GroupId.Get(context);

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authToken, groupId, cancellationToken);
            if (await Task<string>.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);
            string value = json["value"].ToString();
            //if (value.Substring(value.Length - 3).Contains("...")) value = value.Substring(0,value.Length - 3) + "}";

            var values = JsonConvert.DeserializeObject<Dictionary<string, object>>(value.Substring(3,value.Length-6));

            //Prepare output
            string _Id = values["id"].ToString();
            DateTime _createdDateTime = DateTime.Parse(values["createdDateTime"].ToString());
            string _description = values["description"].ToString();
            string _displayName = values["displayName"].ToString();
            string _mail = values["mail"].ToString();
            string _visibility = values["visibility"].ToString();


            // Outputs
            return (ctx) => {
                CreatedDateTime.Set(ctx, _createdDateTime);
                Description.Set(ctx, _description);
                Mail.Set(ctx, _mail);
                Visibillity.Set(ctx, _visibility);
                GroupDictionary.Set(ctx, values);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {

            string restUrl = string.Format("https://graph.microsoft.com/v1.0/groups?$filter=Id eq '{0}'", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }
    }
    #endregion
}

