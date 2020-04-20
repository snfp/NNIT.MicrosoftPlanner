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

namespace NNIT.MicrosoftPlanner.Activities.GroupandUser
{
    [LocalizedDisplayName(nameof(Resources.GetUser_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetUser_Description))]
    public class GetUser : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.GetUser_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_Id_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Id { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_UserPrincipalName_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_UserPrincipalName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> UserPrincipalName { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_BusinessPhones_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_BusinessPhones_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<String[]> BusinessPhones { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_GivenName_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_GivenName_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> GivenName { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_JobTitle_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_JobTitle_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> JobTitle { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_Mail_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_Mail_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Mail { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_MobilePhone_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_MobilePhone_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> MobilePhone { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_OfficeLocation_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_OfficeLocation_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> OfficeLocation { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_PreferredLanguage_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_PreferredLanguage_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> PreferredLanguage { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_Surname_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_Surname_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Surname { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_UserPrincipalName_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_UserPrincipalName_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> OutUserPrincipalName { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_Id_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_Id_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> OutId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetUser_JsonResponse_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetUser_JsonResponse_Description))]
        [LocalizedCategory(nameof(Resources.JsonResponse_Category))]
        public OutArgument<string> JsonResponse { get; set; }

        #endregion


        #region Constructors

        public GetUser()
        {
            Constraints.Add(ActivityConstraints.HasParentType<GetUser, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if ((Id == null && UserPrincipalName == null) || (Id != null && UserPrincipalName != null)) metadata.AddValidationError(string.Format(Resources.ValidationExclusiveProperties_Error, nameof(Id), nameof(UserPrincipalName)));

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
            string userPrincipalName = UserPrincipalName.Get(context);
            Task<string> task;

            // Set a timeout on the execution
            if (id != null) { task = ExecuteWithTimeout(context, authToken, id, cancellationToken); } else { task = ExecuteWithTimeout(context, authToken, userPrincipalName, cancellationToken); }
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            //Wait for task to return with result
            string result = await task;

            //Prepare output
            JObject json = JObject.Parse(result);
            string[] businessPhones = JsonConvert.DeserializeObject<string[]>(json["businessPhones"].ToString());

            // Outputs
            return (ctx) => {
                OutId.Set(ctx, json["id"].ToString());
                BusinessPhones.Set(ctx, businessPhones);
                GivenName.Set(ctx, json["givenName"].ToString());
                JobTitle.Set(ctx, json["jobTitle"].ToString());
                Mail.Set(ctx, json["mail"].ToString());
                MobilePhone.Set(ctx, json["mobilePhone"].ToString());
                OfficeLocation.Set(ctx, json["officeLocation"].ToString());
                PreferredLanguage.Set(ctx, json["preferredLanguage"].ToString());
                Surname.Set(ctx, json["surname"].ToString());
                OutUserPrincipalName.Set(ctx, json["userPrincipalName"].ToString());
                JsonResponse.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string Id, CancellationToken cancellationToken = default)
        {
            string restUrl = string.Format("https://graph.microsoft.com/v1.0/users/{0}",Id);

            HTTPHandler requester = new HTTPHandler();
            return await requester.GetRequest(restUrl, authToken, cancellationToken);
        }

        #endregion
    }
}

