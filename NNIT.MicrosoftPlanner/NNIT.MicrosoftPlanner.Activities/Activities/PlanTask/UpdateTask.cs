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
    [LocalizedDisplayName(nameof(Resources.UpdateTask_DisplayName))]
    [LocalizedDescription(nameof(Resources.UpdateTask_Description))]
    public class UpdateTask : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.UpdateTask_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public InArgument<string> JsonFormat { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_StatusCode_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_StatusCode_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> StatusCode { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_Assignments_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_Assignments_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string,bool>> Assignments { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_AppliedCategories_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_AppliedCategories_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string,Boolean>> AppliedCategories { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_BucketId_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_BucketId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> BucketId { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_DueDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_DueDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<DateTimeOffset> DueDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_PercentComplete_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_PercentComplete_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Nullable<int>> PercentComplete { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_StartDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_StartDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<DateTimeOffset> StartDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_Title_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTask_ETag_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTask_ETag_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ETag { get; set; }

        #endregion


        #region Constructors

        public UpdateTask()
        {
            Constraints.Add(ActivityConstraints.HasParentType<UpdateTask, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Id == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Id)));

            //Fill out json or one of the field in the input category
            if ((Title == null && BucketId == null && StartDateTime == null && DueDateTime == null && PercentComplete == null && Assignments == null && AppliedCategories == null && JsonFormat == null) || ((Title != null || BucketId != null || StartDateTime != null || DueDateTime != null || PercentComplete != null || Assignments != null || AppliedCategories != null) && JsonFormat != null)) 
                metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties2_Error, nameof(Title) + ", " + nameof(BucketId) + ", " + nameof(StartDateTime) + ", " + nameof(DueDateTime) + ", " + nameof(StartDateTime) + ", " + nameof(PercentComplete) + ", " + nameof(Assignments) + ", " + nameof(AppliedCategories), nameof(JsonFormat)));
            
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var id = Id.Get(context);
            var jsonFormat = JsonFormat.Get(context);
            var assignments = Assignments.Get(context);
            var appliedCategories = AppliedCategories.Get(context);
            var bucketId = BucketId.Get(context);
            var duedatetime = DueDateTime.Get(context);
            var percentComplete = PercentComplete.Get(context);
            var startdatetime = StartDateTime.Get(context);
            var title = Title.Get(context);
            var etag = ETag.Get(context);
            string authToken = objectContainer.Get<string>();
            Task<string> task;

            //If no jsonformat, format by using Jsonconvert
            if (string.IsNullOrEmpty(jsonFormat))
            {
                Dictionary<string, object> RequestJson = new Dictionary<string, object>();
                if(!string.IsNullOrEmpty(bucketId)) RequestJson.Add("bucketId", bucketId);
                if (duedatetime != DateTimeOffset.MinValue) RequestJson.Add("dueDateTime", duedatetime);
                if (percentComplete != null) RequestJson.Add("percentComplete", percentComplete);
                if (startdatetime != DateTimeOffset.MinValue) RequestJson.Add("startDateTime", startdatetime);
                if (! string.IsNullOrEmpty(title)) RequestJson.Add("title", title);
                if (appliedCategories != null) RequestJson.Add("appliedCategories", appliedCategories);
                if (assignments != null)
                {
                    Dictionary<string, Dictionary<string,string>> assigment = new Dictionary<string, Dictionary<string, string>>();

                    foreach (string user in assignments.Keys)
                    {
                        if(assignments[user])
                        assigment.Add(user, new Dictionary<string, string>() { { "@odata.type", "#microsoft.graph.plannerAssignment" }, { "orderHint", " !" } });
                        else
                        {
                            assigment.Add(user, null);
                        }
                    }
                    RequestJson.Add("assignments", assigment);
                }
                jsonFormat = JsonConvert.SerializeObject(RequestJson);
            }

            // Set a timeout on the execution
            if (etag != null) { task = ExecuteWithTimeout(context, authToken, id,etag, jsonFormat, cancellationToken); } else { task = ExecuteWithTimeout(context, authToken, id, jsonFormat, cancellationToken); }
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            // Outputs
            return (ctx) => {
                StatusCode.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, string jsonInput, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}", id);

            //Get etag 
            HTTPHandler requester = new HTTPHandler();
            string Result = await requester.GetRequest(restUrl, authToken, cancellationToken);
            JObject json = JObject.Parse(Result);
            string etag = json["@odata.etag"].ToString();
            //Use etag to delete 
            return await ExecuteWithTimeout(context, authToken, id, etag, jsonInput, cancellationToken);
        }
        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, string etag, string jsonInput, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.PatchRequest(restUrl, authToken, etag, jsonInput, cancellationToken);
        }

        #endregion
    }
}

