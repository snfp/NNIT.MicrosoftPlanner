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

namespace NNIT.MicrosoftPlanner.Activities.Plan
{

    [LocalizedDisplayName(nameof(Resources.UpdatePlan_DisplayName))]
    [LocalizedDescription(nameof(Resources.UpdatePlan_Description))]
    public class UpdatePlan : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.UpdatePlan_PlanId_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlan_PlanId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> PlanId { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlan_ETag_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlan_ETag_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ETag { get; set; }


        [LocalizedDisplayName(nameof(Resources.UpdatePlan_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlan_Title_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlan_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlan_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public InArgument<string> JsonFormat { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdatePlan_StatusCode_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdatePlan_StatusCode_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> StatusCode { get; set; }

        #endregion


        #region Constructors

        public UpdatePlan()
        {
            Constraints.Add(ActivityConstraints.HasParentType<UpdatePlan, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (PlanId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Id)));

            if ((Title == null && JsonFormat == null) || (Title != null && JsonFormat != null)) metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties_Error, nameof(Title),nameof(JsonFormat)));
           
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var id = PlanId.Get(context);
            var etag = ETag.Get(context);
            var title = Title.Get(context);
            var jsonFormat = JsonFormat.Get(context);
            string authToken = objectContainer.Get<string>();
            Task<string> task;

            //If no jsonformat was provided, format it by using JObject
            if (string.IsNullOrEmpty(jsonFormat))
            {
                JObject jsonObject =
                          new JObject(
                            new JProperty("title", title));
                jsonFormat = jsonObject.ToString();
            }

            // Set a timeout on the execution
            if (etag != null) { task = ExecuteWithTimeout(context, authToken, id, etag, jsonFormat, cancellationToken); } else { task = ExecuteWithTimeout(context, authToken, id, jsonFormat, cancellationToken); }
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
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}", "oTXRrczdIkqkTjOHCDuwo5YAFKpq");

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

            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}", "oTXRrczdIkqkTjOHCDuwo5YAFKpq");

            HTTPHandler requester = new HTTPHandler();
            return await requester.PatchRequest(restUrl, authToken, etag, jsonInput, cancellationToken);
        }
        #endregion
    }
}

