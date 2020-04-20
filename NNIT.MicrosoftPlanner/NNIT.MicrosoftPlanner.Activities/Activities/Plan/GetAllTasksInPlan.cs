using System;
using System.Activities;
using System.Threading;
using NNIT.MicrosoftPlanner.Activities.Properties;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;
using UiPath.Shared.Activities.HTTP;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace NNIT.MicrosoftPlanner.Activities.Plan
{
    [LocalizedDisplayName(nameof(Resources.GetAllTasksInPlan_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetAllTasksInPlan_Description))]
    public class GetAllTasksInPlan : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetAllTasksInPlan_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllTasksInPlan_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetAllTasksInPlan_Tasks_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllTasksInPlan_Tasks_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<List<Dictionary<string,object>>> Tasks { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetAllTasksInPlan_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetAllTasksInPlan_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetAllTasksInPlan()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetAllTasksInPlan, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
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
            string authToken = objectContainer.Get<string>();
            string id = Id.Get(context);

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context,authToken, id, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);
            int NumberOfPlands = json["value"].Count();

            List<Dictionary<string, object>> alltasks = new List<Dictionary<string, object>>();

            //Go through each plan and extract relavant information
            for (int i = 0; NumberOfPlands > i; i++)
            {
                Dictionary<string, object> singleTask = new Dictionary<string, object>();
                var value = json["value"][i].ToString();
                var values = JsonConvert.DeserializeObject<Dictionary<string, object>>(value);

                foreach (string key in values.Keys)
                {
                    if (key == "createdBy")
                    {
                        string innerId = json["value"][i][key]["user"]["id"].ToString();
                        singleTask.Add(key, innerId);
                    }
                    else if (key == "appliedCategories")
                    {
                        string innerLevel = json["value"][i][key].ToString();
                        Dictionary<string, Boolean> appliedCategories = JsonConvert.DeserializeObject<Dictionary<string, Boolean>>(innerLevel);
                        singleTask.Add(key, appliedCategories);
                    }
                    else if (key == "assignments")
                    {
                        string innerLevel = json["value"][i][key].ToString();
                        Dictionary<string, object> assignments = JsonConvert.DeserializeObject<Dictionary<string, object>>(innerLevel);
                        string[] assigned = assignments.Keys.ToArray();
                        singleTask.Add(key, assigned);
                    }
                    else
                    {
                        singleTask.Add(key, values[key]);
                    }

                }
                alltasks.Add(singleTask);
            }

            // Outputs
            return (ctx) => {
                Tasks.Set(ctx, alltasks);
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/planner/plans/{0}/tasks", id); ;

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

