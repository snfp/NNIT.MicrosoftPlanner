using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using NNIT.MicrosoftPlanner.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities.Utilities;
using UiPath.Shared.Activities.HTTP;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Linq;

namespace NNIT.MicrosoftPlanner.Activities.GroupandUser
{
    [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_Description))]
    public class GetPlansRelatedToGroupId : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_GroupId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_GroupId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> GroupId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_PlanDictionaries_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_PlanDictionaries_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<List<Dictionary<string, object>>> PlanDictionaries { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_Owners_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_Owners_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string[]> Owners { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_Ids_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_Ids_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string[]> Ids { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_Titles_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_Titles_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string[]> Titles { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlansRelatedToGroupId_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlansRelatedToGroupId_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        private string _groupId;
        private string _authtoken;

        #endregion


        #region Constructors

        public GetPlansRelatedToGroupId()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetPlansRelatedToGroupId, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            var timeout = TimeoutMS.Get(context);
            _groupId = GroupId.Get(context);
            _authtoken = objectContainer.Get<string>();

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare Output
            JObject json = JObject.Parse(result);
            int NumberOfPlans = json["value"].Count();
            List<Dictionary<string, object>> plans = new List<Dictionary<string, object>>();
            string[] owners = new string[NumberOfPlans];
            string[] ids = new string[NumberOfPlans];
            string[] DisplayNames = new string[NumberOfPlans];

            for (int i = 0; NumberOfPlans > i; i++)
            {
                var value = json["value"][i].ToString();
                var values = JsonConvert.DeserializeObject<Dictionary<string, object>>(value);
                owners[i] = values["owner"].ToString();
                DisplayNames[i] = values["title"].ToString();
                ids[i] = values["id"].ToString();
                plans.Add(values);
            }

            // Outputs
            return (ctx) => {
                PlanDictionaries.Set(ctx, plans);
                Owners.Set(ctx, owners);
                Ids.Set(ctx, ids);
                Titles.Set(ctx, DisplayNames);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/groups/{0}/planner/plans", _groupId);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, _authtoken,cancellationToken);

        }

        #endregion
    }
}

