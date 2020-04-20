using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NNIT.MicrosoftPlanner.Activities.Properties;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.HTTP;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;

namespace NNIT.MicrosoftPlanner.Activities.PlanTask
{
    [LocalizedDisplayName(nameof(Resources.CreateTask_DisplayName))]
    [LocalizedDescription(nameof(Resources.CreateTask_Description))]
    public class CreateTask : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.CreateTask_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.CreateTask_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> JsonFormat { get; set; }

        [LocalizedDisplayName(nameof(Resources.CreateTask_PlanId_DisplayName))]
        [LocalizedDescription(nameof(Resources.CreateTask_PlanId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> PlanId { get; set; }

        [LocalizedDisplayName(nameof(Resources.CreateTask_BucketId_DisplayName))]
        [LocalizedDescription(nameof(Resources.CreateTask_BucketId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> BucketId { get; set; }

        [LocalizedDisplayName(nameof(Resources.CreateTask_AssignedUsers_DisplayName))]
        [LocalizedDescription(nameof(Resources.CreateTask_AssignedUsers_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<String[]> AssignedUsers { get; set; }

        [LocalizedDisplayName(nameof(Resources.CreateTask_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.CreateTask_Title_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.Createbucket_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        [LocalizedDisplayName(nameof(Resources.Createbucket_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.Createbucket_Id_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Id { get; set; }

        #endregion


        #region Constructors

        public CreateTask()
        {
            Constraints.Add(ActivityConstraints.HasParentType<CreateTask, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            //Fill out json or one of the field in the input category
            if ((PlanId == null && BucketId == null && Title == null && AssignedUsers == null && JsonFormat == null) || ((PlanId != null || BucketId != null || Title != null || AssignedUsers != null) && JsonFormat != null))
                metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties2_Error, nameof(PlanId) + ", " + nameof(BucketId) + ", " + nameof(Title) + ", " + nameof(AssignedUsers), nameof(JsonFormat)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            string jsonformat = JsonFormat.Get(context);
            var planid = PlanId.Get(context);
            var bucketid = BucketId.Get(context);
            string[] assignedUsers = AssignedUsers.Get(context);
            var title = Title.Get(context);
            string authToken = objectContainer.Get<string>();


            //if no Jsonformat provided use JsonConvert to Format json
            if (string.IsNullOrEmpty(jsonformat))
            {
                Dictionary<string, object> RequestJson = new Dictionary<string, object>()
                {
                {"planId", planid},
                {"bucketId", bucketid },
                {"title", title }
                }; ;
                if(assignedUsers != null)
                {
                    Dictionary<string, Dictionary<string, string>> assignments = new Dictionary<string, Dictionary<string, string>>();
                    Dictionary<string, string> mandatoryFillvalues = new Dictionary<string, string>()
                    {
                    { "@odata.type", "#microsoft.graph.plannerAssignment" },
                    {"orderHint", " !" }
                    }; ;
                    foreach (string User in assignedUsers) { assignments.Add(User, mandatoryFillvalues); }

                    RequestJson.Add("assignments", assignments);
                }
               jsonformat = JsonConvert.SerializeObject(RequestJson);
            }
            // Set a timeout on the execution
            var task = ExecuteWithTimeout(context, authToken, jsonformat, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;
            JObject json = JObject.Parse(result);
            Console.WriteLine("Task was created with id " + json["id"].ToString());

            // Outputs
            return (ctx) => {
                JsonResponse.Set(ctx, result);
                Id.Set(ctx, json["id"].ToString());
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string json, CancellationToken cancellationToken = default)
        {
            string restUrl = "https://graph.microsoft.com/v1.0/planner/tasks";

            //Get etag 
            HTTPHandler requester = new HTTPHandler();

            return await requester.PostRequest(restUrl, authToken, json, cancellationToken);
        }

        #endregion
    }
}

