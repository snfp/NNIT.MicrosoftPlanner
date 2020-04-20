using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using NNIT.MicrosoftPlanner.Activities.Properties;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.HTTP;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;

namespace NNIT.MicrosoftPlanner.Activities
{
    [LocalizedDisplayName(nameof(Resources.Createbucket_DisplayName))]
    [LocalizedDescription(nameof(Resources.Createbucket_Description))]
    public class Createbucket : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.Createbucket_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public InArgument<string> JsonFormat { get; set; } = null;

        [LocalizedDisplayName(nameof(Resources.Createbucket_PlanId_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_PlanId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> PlanId { get; set; }

        [LocalizedDisplayName(nameof(Resources.Createbucket_BucketName_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_BucketName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> BucketName { get; set; }

        [LocalizedDisplayName(nameof(Resources.Createbucket_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        [LocalizedDisplayName(nameof(Resources.Createbucket_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_Id_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Id { get; set; }
        #endregion


        #region Constructors

        public Createbucket()
        {
            Constraints.Add(ActivityConstraints.HasParentType<Createbucket, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (((BucketName != null || PlanId != null) && JsonFormat != null)) metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties_Error, nameof(BucketName) + " and " + nameof(PlanId), nameof(JsonFormat)));
            else if ((BucketName == null || PlanId == null) && JsonFormat == null) metadata.AddValidationError(string.Format(Resources.ValidationOverloadGroup_Error, nameof(BucketName), nameof(PlanId)));
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);
            // Inputs
            var timeout = TimeoutMS.Get(context);
            var jsonformat = JsonFormat.Get(context);
            var planid = PlanId.Get(context);
            var bucketname = BucketName.Get(context);
            string authToken = objectContainer.Get<string>();
            
            //if no jsonformat, format the json by using Jobject
            if (string.IsNullOrEmpty(jsonformat))
            {
                JObject jsonObject =
                          new JObject(
                            new JProperty("name", bucketname),
                            new JProperty("planId", planid),
                            new JProperty("orderHint", " !"));
                jsonformat = jsonObject.ToString();
            }
                // Set a timeout on the execution
                Task<string> task = ExecuteWithTimeout(context, authToken, jsonformat, cancellationToken);
                if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);


                //Wait for task to return with result
                string result = await task;

                JObject json = JObject.Parse(result);
                Console.WriteLine("Bucket was created with id: " + json["id"].ToString());
                // Outputs
                return (ctx) =>
                {
                    JsonResponse.Set(ctx, result);
                    Id.Set(ctx, json["id"].ToString());
                };
        }


        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string json, CancellationToken cancellationToken = default)
        {
            string restUrl = "https://graph.microsoft.com/v1.0/planner/buckets";

            //Get etag 
            HTTPHandler requester = new HTTPHandler();

            return await requester.PostRequest(restUrl, authToken, json, cancellationToken);
        }
        #endregion
    }
}

