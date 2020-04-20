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
using Newtonsoft.Json.Linq;
using UiPath.Shared.Activities.HTTP;
using Newtonsoft.Json;
using System.Linq;

namespace NNIT.MicrosoftPlanner.Activities.PlanTask
{
    [LocalizedDisplayName(nameof(Resources.GetTask_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetTask_Description))]
    public class GetTask : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetTask_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_PlanId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_PlanId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> PlanId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_BucketId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_BucketId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> BucketId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_Title_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_PercentComplete_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_PercentComplete_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Int32> PercentComplete { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_StartDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_StartDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> StartDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_CreatedDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_CreatedDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> CreatedDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_DueDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_DueDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> DueDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_HasDescription_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_HasDescription_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Boolean> HasDescription { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_CompletedDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_CompletedDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> CompletedDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_CompletedBy_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_CompletedBy_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> CompletedBy { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_CreatedBy_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_CreatedBy_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> CreatedBy { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_ReferenceCount_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_ReferenceCount_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<int> ReferenceCount { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_ChecklistItemCount_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_ChecklistItemCount_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<int> ChecklistItemCount { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_AppliedCategories_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_AppliedCategories_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, string>> AppliedCategories { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_AssignedUsers_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_AssignedUsers_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string[]> AssignedUsers { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_ConversationThreadId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_ConversationThreadId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> ConversationThreadId { get; set; }


        [LocalizedDisplayName(nameof(Resources.GetTask_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTask_TaskAsDictionary_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTask_TaskAsDictionary_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, object>> TaskAsDictionary { get; set; }

        #endregion


        #region Constructors

        public GetTask()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetTask, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            string authToken = objectContainer.Get<string>();
            var id = Id.Get(context);

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authToken, id, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);
            Dictionary<string, object> taskDictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json.ToString());
            Dictionary<string, string> appliedCategories = JsonConvert.DeserializeObject<Dictionary<string, string>>(json["appliedCategories"].ToString());
            string[] assignments = JsonConvert.DeserializeObject<Dictionary<string, object>>(json["assignments"].ToString()).Keys.ToArray();
            taskDictionary["assignments"] = assignments;
            taskDictionary["appliedCategories"] = appliedCategories;
            taskDictionary["createdBy"] = json["createdBy"]["user"]["id"].ToString();

            DateTime startDateTime;
            DateTime dueDateTime;
            DateTime completedDateTime;

            // Outputs
            return (ctx) => {
                Etag.Set(ctx, json["@odata.etag"].ToString());
                PlanId.Set(ctx, json["planId"].ToString());
                BucketId.Set(ctx, json["bucketId"].ToString());
                Title.Set(ctx, json["title"].ToString());
                PercentComplete.Set(ctx, int.Parse(json["percentComplete"].ToString()));
                if (DateTime.TryParse(json["startDateTime"].ToString(), out startDateTime)) 
                    StartDateTime.Set(ctx, startDateTime);
                if (DateTime.TryParse(json["dueDateTime"].ToString(), out dueDateTime))
                    DueDateTime.Set(ctx, dueDateTime);
                CreatedDateTime.Set(ctx, DateTime.Parse(json["createdDateTime"].ToString()));
                HasDescription.Set(ctx, bool.Parse(json["hasDescription"].ToString()));
                if (DateTime.TryParse(json["completedDateTime"].ToString(), out completedDateTime))
                    CompletedDateTime.Set(ctx, completedDateTime);
                CompletedBy.Set(ctx, json["completedBy"].ToString());
                CreatedBy.Set(ctx, json["createdBy"]["user"]["id"].ToString());
                ReferenceCount.Set(ctx, int.Parse(json["referenceCount"].ToString()));
                ChecklistItemCount.Set(ctx, int.Parse(json["checklistItemCount"].ToString()));
                AppliedCategories.Set(ctx, appliedCategories);
                AssignedUsers.Set(ctx, assignments);
                ConversationThreadId.Set(ctx, json["conversationThreadId"].ToString());
                JsonResponse.Set(ctx, result);
                TaskAsDictionary.Set(ctx, taskDictionary);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

