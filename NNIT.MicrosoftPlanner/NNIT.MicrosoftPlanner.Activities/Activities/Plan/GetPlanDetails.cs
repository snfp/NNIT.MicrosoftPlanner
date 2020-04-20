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
using UiPath.Shared.Activities.HTTP;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace NNIT.MicrosoftPlanner.Activities.Plan
{
    [LocalizedDisplayName(nameof(Resources.GetPlanDetails_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetPlanDetails_Description))]
    public class GetPlanDetails : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetPlanDetails_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlanDetails_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlanDetails_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlanDetails_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlanDetails_SharedWith_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlanDetails_SharedWith_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, bool>> SharedWith { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlanDetails_CategoryDescriptions_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlanDetails_CategoryDescriptions_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, string>> CategoryDescriptions { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlanDetails_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlanDetails_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetPlanDetails()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetPlanDetails, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Id == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Id)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {

            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var id = Id.Get(context);
            string authToken = objectContainer.Get<string>();

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authToken, id, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);


            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);
            Dictionary<string, Boolean>  sharedWith = JsonConvert.DeserializeObject<Dictionary<string, Boolean>>(json["sharedWith"].ToString());
            Dictionary<string, string> categoryDescriptions = JsonConvert.DeserializeObject<Dictionary<string, string>>(json["categoryDescriptions"].ToString());

            // Outputs
            return (ctx) => {
                Etag.Set(ctx, json["@odata.etag"].ToString());
                SharedWith.Set(ctx, sharedWith);
                CategoryDescriptions.Set(ctx, categoryDescriptions);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}/details", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

