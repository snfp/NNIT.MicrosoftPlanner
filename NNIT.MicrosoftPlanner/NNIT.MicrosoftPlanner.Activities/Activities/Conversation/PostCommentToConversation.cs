using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using NNIT.MicrosoftPlanner.Activities.Properties;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.HTTP;
using UiPath.Shared.Activities.Localization;
using UiPath.Shared.Activities.Utilities;

namespace NNIT.MicrosoftPlanner.Activities.Conversation
{
    [LocalizedDisplayName(nameof(Resources.PostCommentToConversation_DisplayName))]
    [LocalizedDescription(nameof(Resources.PostCommentToConversation_Description))]
    public class PostCommentToConversation : ContinuableAsyncCodeActivity
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

        [LocalizedDisplayName(nameof(Resources.PostCommentToConversation_GroupId_DisplayName))]
        [LocalizedDescription(nameof(Resources.PostCommentToConversation_GroupId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> GroupId { get; set; }

        [LocalizedDisplayName(nameof(Resources.PostCommentToConversation_ConversationTreadId_DisplayName))]
        [LocalizedDescription(nameof(Resources.PostCommentToConversation_ConversationTreadId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ConversationTreadId { get; set; }

        [LocalizedDisplayName(nameof(Resources.PostCommentToConversation_Comment_DisplayName))]
        [LocalizedDescription(nameof(Resources.PostCommentToConversation_Comment_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Comment { get; set; }

        [LocalizedDisplayName(nameof(Resources.PostCommentToConversation_StatusCode_DisplayName))]
        [LocalizedDescription(nameof(Resources.PostCommentToConversation_StatusCode_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> StatusCode { get; set; }

        #endregion


        #region Constructors

        public PostCommentToConversation()
        {
            Constraints.Add(ActivityConstraints.HasParentType<PostCommentToConversation, PlannerScope>(string.Format(Resources.ValidationScope_Error, Resources.PlannerScope_DisplayName)));
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (GroupId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(GroupId)));
            if (ConversationTreadId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ConversationTreadId)));
            if (Comment == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Comment)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {

            // Object Container: Use objectContainer.Get<T>() to retrieve objects from the scope
            var objectContainer = context.GetFromContext<IObjectContainer>(PlannerScope.ParentContainerPropertyTag);

            // Inputs
            var timeout = TimeoutMS.Get(context);
            var groupId = GroupId.Get(context);
            var conversationtreadid = ConversationTreadId.Get(context);
            string authToken = objectContainer.Get<string>();
            var comment = Comment.Get(context);

            //Generate json
            string jsonformat = "{\"post\": {\"body\": {\"contentType\": \"1\",\"content\": \"" + comment + "\"}}}";

            // Set a timeout on the execution
            Task<string> task = ExecuteWithTimeout(context, authToken, groupId, conversationtreadid, jsonformat, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);
            string result = await task;


            // Outputs
            return (ctx) => {
                StatusCode.Set(ctx, result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, string authToken, string groupId, string conversationId, string jsonInput,CancellationToken cancellationToken = default)
        {

            string restUrl = string.Format("https://graph.microsoft.com/v1.0/groups/{0}/threads/{1}/reply", groupId, conversationId);

            HTTPHandler requester = new HTTPHandler();
            return await requester.PostRequest(restUrl, authToken, jsonInput, cancellationToken);
        }
        #endregion
    }
}

