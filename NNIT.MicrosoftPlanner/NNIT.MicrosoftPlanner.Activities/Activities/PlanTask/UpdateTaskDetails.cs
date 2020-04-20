using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NNIT.MicrosoftPlanner.Activities.Properties;
using NNIT.MircrosoftPlanner.Enums;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.HTTP;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;

namespace NNIT.MicrosoftPlanner.Activities.PlanTask
{
    [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_DisplayName))]
    [LocalizedDescription(nameof(Resources.UpdateTaskDetails_Description))]
    public class UpdateTaskDetails : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_Description_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_Description_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Description { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_JsonFormat_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_JsonFormat_Description))]
        [LocalizedCategory(nameof(Resources.JsonInput_Category))]
        public InArgument<string> JsonFormat { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_PreviewType_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_PreviewType_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        [TypeConverter(typeof(EnumNameConverter<PreviewTypes>))]
        public PreviewTypes PreviewType { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_CheckList_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_CheckList_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string,Dictionary<string,object>>> CheckList { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_References_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_References_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<Dictionary<string, Dictionary<string, string>>> References { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_Etag_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_Etag_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Etag { get; set; }

        [LocalizedDisplayName(nameof(Resources.UpdateTaskDetails_StatusCode_DisplayName))]
        [LocalizedDescription(nameof(Resources.UpdateTaskDetails_StatusCode_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> StatusCode { get; set; }

        #endregion


        #region Constructors

        public UpdateTaskDetails()
        {
            Constraints.Add(ActivityConstraints.HasParentType<UpdateTaskDetails, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Id == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Id)));

            //Fill out json or one of the field in the input category
            if ((Description == null && PreviewType == PreviewTypes.NoChange && CheckList == null && References == null && JsonFormat == null) || ((Description != null || PreviewType != PreviewTypes.NoChange || CheckList != null || References != null) && JsonFormat != null))
                metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties2_Error, nameof(Description) + ", " + nameof(PreviewType) + ", " + nameof(CheckList) + ", " + nameof(References), nameof(JsonFormat)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var id = Id.Get(context);
            var description = Description.Get(context);
            var jsonFormat = JsonFormat.Get(context);
            var previewtype = PreviewType;
            var checklist = CheckList.Get(context);
            var references = References.Get(context);
            var etag = Etag.Get(context);
            string authToken = objectContainer.Get<string>();
            Task<string> task;

            //If no jsonformat, format by using JsonConvert
            if (string.IsNullOrEmpty(jsonFormat))
            {
                Dictionary<string, object> RequestJson = new Dictionary<string, object>();
                if (! string.IsNullOrEmpty(description)) RequestJson.Add("description", description);
                if (previewtype != PreviewTypes.NoChange) RequestJson.Add("previewType", previewtype.ToString());
                if (checklist != null) RequestJson.Add("checklist", checklist);
                if (references != null) RequestJson.Add("references", references);
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
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}/details", id);

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
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/tasks/{0}/details", id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.PatchRequest(restUrl, authToken, etag, jsonInput, cancellationToken);
        }

        #endregion
    }
}

