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
    [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_DisplayName))]
    [LocalizedDescription(nameof(Resources.UpdatePlanDetails_Description))]
    public class UpdatePlanDetails : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_SharedWith_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_SharedWith_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string, bool>> SharedWith { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_CategoryDescriptions_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_CategoryDescriptions_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string, string>> CategoryDescriptions { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public InArgument<string> JsonFormat { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlanDetails_StatusCode_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlanDetails_StatusCode_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> StatusCode { get; set; }

        #endregion


        #region Constructors

        public UpdatePlanDetails()
        {
            Constraints.Add(ActivityConstraints.HasParentType<UpdatePlanDetails, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Id == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Id)));
            if ((CategoryDescriptions == null && SharedWith == null && JsonFormat== null) || ((CategoryDescriptions != null || SharedWith != null) && JsonFormat != null)) metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties_Error, nameof(CategoryDescriptions)+ " and " + nameof(SharedWith), nameof(JsonFormat)));


            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var id = Id.Get(context);
            var etag = Etag.Get(context);
            var sharedWith = SharedWith.Get(context);
            var categoryDescriptions = CategoryDescriptions.Get(context);
            var jsonFormat = JsonFormat.Get(context);
            string authToken = objectContainer.Get<string>();
            Task<string> task;

            //If no jsonformat was provided format json by using JsonConvert
            if (string.IsNullOrEmpty(jsonFormat))
            {
                Dictionary<string, object> RequestJson = new Dictionary<string, object>();
                if(sharedWith != null)RequestJson.Add("sharedWith", sharedWith);
                if (categoryDescriptions != null)RequestJson.Add("categoryDescriptions", categoryDescriptions);
                jsonFormat = JsonConvert.SerializeObject(RequestJson);
            }


            // Set a timeout on the execution
            if (etag != null) { task = ExecuteWithTimeout(context, authToken, id, etag, jsonFormat, cancellationToken); } else { task = ExecuteWithTimeout(context, authToken, id, jsonFormat, cancellationToken); }
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);


            //Wait for task to return with result
            string result = await task;

            // Outputs
            return (ctx) => {
                StatusCode.Set(ctx, task.Result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, string jsonInput, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}/details", "oTXRrczdIkqkTjOHCDuwo5YAFKpq");

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

            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}/details", "oTXRrczdIkqkTjOHCDuwo5YAFKpq");

            HTTPHandler requester = new HTTPHandler();
            return await requester.PatchRequest(restUrl, authToken, etag, jsonInput, cancellationToken);
        }

        #endregion
    }
}

