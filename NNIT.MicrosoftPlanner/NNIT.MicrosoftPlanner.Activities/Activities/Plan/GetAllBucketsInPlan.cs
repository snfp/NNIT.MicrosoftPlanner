using System;
using System.Activities;
using System.Collections.Generic;
using NNIT.MicrosoftPlanner.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;
using NNIT.MircrosoftPlanner.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPath.Shared.Activities.HTTP;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace NNIT.MicrosoftPlanner.Activities.Plan
{
    [LocalizedDisplayName(nameof(Resources.GetAllBucketsInPlan_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetAllBucketsInPlan_Description))]
    public class GetAllBucketsInPlan : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetAllBucketsInPlan_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllBucketsInPlan_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetAllBucketsInPlan_Buckets_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllBucketsInPlan_Buckets_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<List<Dictionary<string, string>>> Buckets { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetAllBucketsInPlan_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllBucketsInPlan_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetAllBucketsInPlan()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetAllBucketsInPlan, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            int timeout = TimeoutMS.Get(context);
            string id = Id.Get(context);
            string authtoken = objectContainer.Get<string>();

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authtoken, id, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare output 
            JObject json = JObject.Parse(result);
            List<Dictionary<string, string>> buckets = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(json["value"].ToString());
           
            // Outputs
            return (ctx) => {
                Buckets.Set(ctx, buckets);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}/buckets", id); ;

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

