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

namespace NNIT.MicrosoftPlanner.Activities.Plan
{
    [LocalizedDisplayName(nameof(Resources.GetPlan_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetPlan_Description))]
    public class GetPlan : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetPlan_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_CreatedDateTime_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_CreatedDateTime_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DateTime> CreatedDateTime { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_Owner_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_Owner_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Owner { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_Title_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_CreatedByUser_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_CreatedByUser_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> CreatedByUser { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_CreatedByApplication_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_CreatedByApplication_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> CreatedByApplication { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetPlan_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetPlan_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetPlan()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetPlan, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            string authToken = objectContainer.Get<string>();

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authToken, id, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

    
            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);

            // Outputs
            return (ctx) => {
                Etag.Set(ctx, json["@odata.etag"].ToString());
                CreatedDateTime.Set(ctx, DateTime.Parse(json["createdDateTime"].ToString()));
                Owner.Set(ctx, json["owner"].ToString());
                Title.Set(ctx, json["title"].ToString());
                CreatedByUser.Set(ctx, json["createdBy"]["user"]["id"].ToString());
                CreatedByUser.Set(ctx, json["createdBy"]["application"]["id"].ToString());
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

