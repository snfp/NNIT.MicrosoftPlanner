using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using NNIT.MicrosoftPlanner.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using NNIT.MircrosoftPlanner.Activities;
using NNIT.MicrosoftPlanner.Activities.Plan;
using UiPath.Shared.Activities.Utilities;
using UiPath.Shared.Activities.HTTP;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace NNIT.MicrosoftPlanner.Activities.PlanBucket
{
    [LocalizedDisplayName(nameof(Resources.GetBucket_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetBucket_Description))]
    public class GetBucket : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetBucket_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetBucket_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetBucket_PlanId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_PlanId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> PlanId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetBucket_Name_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_Name_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Name { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetBucket_JsonRepsonse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_JsonRepsonse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonRepsonse { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetBucket_BucketDictionary_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetBucket_BucketDictionary_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, string>> BucketDictionary { get; set; }

        #endregion


        #region Constructors

        public GetBucket()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetBucket, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            Dictionary<string, string> bucket = JsonConvert.DeserializeObject<Dictionary<string, string>>(json.ToString());

            // Outputs
            return (ctx) => {
                Etag.Set(ctx, json["@odata.etag"].ToString());
                PlanId.Set(ctx, json["planId"].ToString());
                Name.Set(ctx, json["name"].ToString());
                JsonRepsonse.Set(ctx, result);
                BucketDictionary.Set(ctx, bucket);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/buckets/{0}", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

