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

namespace NNIT.MicrosoftPlanner.Activities.PlanTask
{
    [LocalizedDisplayName(nameof(Resources.GetTaskDetails_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetTaskDetails_Description))]
    public class GetTaskDetails : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_Description_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_Description_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Description { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_Reference_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_Reference_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, Dictionary<string, object>>> Reference { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_CheckList_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_CheckList_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, Dictionary<string, object>>> CheckList { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_TaskDictionary_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_TaskDictionary_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<Dictionary<string, object>> TaskDictionary { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetTaskDetails_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetTaskDetails_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetTaskDetails()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetTaskDetails, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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

            Dictionary<string, Dictionary<string, object>> references = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(json["references"].ToString());
            foreach (string key in references.Keys) { references[key]["lastModifiedBy"] = json["references"][key]["lastModifiedBy"]["user"]["id"]; }
            taskDictionary["references"] = references;

            Dictionary<string, Dictionary<string, object>> checklist = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(json["checklist"].ToString());
            foreach (string key in checklist.Keys) { checklist[key]["lastModifiedBy"] = json["checklist"][key]["lastModifiedBy"]["user"]["id"]; }
            taskDictionary["checklist"] = checklist;

            // Outputs
            return (ctx) => {
                Description.Set(ctx, json["description"].ToString());
                Etag.Set(ctx, json["@odata.etag"].ToString());
                Reference.Set(ctx, references);
                CheckList.Set(ctx, checklist);
                TaskDictionary.Set(ctx, taskDictionary);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id,  CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}/details", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

