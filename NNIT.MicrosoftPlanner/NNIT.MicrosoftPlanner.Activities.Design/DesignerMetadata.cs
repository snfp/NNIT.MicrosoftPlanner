using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using NNIT.MicrosoftPlanner.Activities.Design.Designers;
using NNIT.MicrosoftPlanner.Activities.Design.Properties;
using NNIT.MicrosoftPlanner.Activities.GroupandUser;
using NNIT.MicrosoftPlanner.Activities.Plan;
using NNIT.MicrosoftPlanner.Activities.PlanBucket;
using NNIT.MicrosoftPlanner.Activities.PlanTask;
using NNIT.MicrosoftPlanner.Activities.Conversation;
using NNIT.MircrosoftPlanner.Activities;

namespace NNIT.MicrosoftPlanner.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            //Setup Categories
            var RootCategory = new CategoryAttribute($"{Resources.Category}");
            var GroupAndUsersCategory = new CategoryAttribute($"{Resources.Category_GroupAndUser}");
            var PlanCategory = new CategoryAttribute($"{Resources.Category_Plan}");
            var BucketCategory = new CategoryAttribute($"{Resources.Category_Bucket}");
            var TaskCategory = new CategoryAttribute($"{Resources.Category_Task}");
            var ConversationCategory = new CategoryAttribute($"{Resources.Category_Conversation}");

            builder.AddCustomAttributes(typeof(PlannerScope), RootCategory);
            builder.AddCustomAttributes(typeof(PlannerScope), new DesignerAttribute(typeof(PlannerScopeDesigner)));
            builder.AddCustomAttributes(typeof(PlannerScope), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetPlannerGroup), PlanCategory);
            builder.AddCustomAttributes(typeof(GetPlannerGroup), new DesignerAttribute(typeof(GetPlannerGroupDesigner)));
            builder.AddCustomAttributes(typeof(GetPlannerGroup), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetPlansRelatedToGroupId), GroupAndUsersCategory);
            builder.AddCustomAttributes(typeof(GetPlansRelatedToGroupId), new DesignerAttribute(typeof(GetPlansRelatedToGroupIdDesigner)));
            builder.AddCustomAttributes(typeof(GetPlansRelatedToGroupId), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetPlanDetails), PlanCategory);
            builder.AddCustomAttributes(typeof(GetPlanDetails), new DesignerAttribute(typeof(GetPlanDetailsDesigner)));
            builder.AddCustomAttributes(typeof(GetPlanDetails), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetPlan), PlanCategory);
            builder.AddCustomAttributes(typeof(GetPlan), new DesignerAttribute(typeof(GetPlanDetailsDesigner)));
            builder.AddCustomAttributes(typeof(GetPlan), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetUser), GroupAndUsersCategory);
            builder.AddCustomAttributes(typeof(GetUser), new DesignerAttribute(typeof(GetUserDesigner)));
            builder.AddCustomAttributes(typeof(GetUser), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetAllBucketsInPlan), PlanCategory);
            builder.AddCustomAttributes(typeof(GetAllBucketsInPlan), new DesignerAttribute(typeof(GetAllBucketsInPlanDesigner)));
            builder.AddCustomAttributes(typeof(GetAllBucketsInPlan), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetAllTasksInPlan), PlanCategory);
            builder.AddCustomAttributes(typeof(GetAllTasksInPlan), new DesignerAttribute(typeof(GetAllTasksInPlanDesigner)));
            builder.AddCustomAttributes(typeof(GetAllTasksInPlan), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetAllMyTasks), TaskCategory);
            builder.AddCustomAttributes(typeof(GetAllMyTasks), new DesignerAttribute(typeof(GetAllMyTasksDesigner)));
            builder.AddCustomAttributes(typeof(GetAllMyTasks), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetTask), TaskCategory);
            builder.AddCustomAttributes(typeof(GetTask), new DesignerAttribute(typeof(GetTaskDesigner)));
            builder.AddCustomAttributes(typeof(GetTask), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetTaskDetails), TaskCategory);
            builder.AddCustomAttributes(typeof(GetTaskDetails), new DesignerAttribute(typeof(GetTaskDetailsDesigner)));
            builder.AddCustomAttributes(typeof(GetTaskDetails), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetAllTasksinBucket), BucketCategory);
            builder.AddCustomAttributes(typeof(GetAllTasksinBucket), new DesignerAttribute(typeof(GetAllTasksinBucketDesigner)));
            builder.AddCustomAttributes(typeof(GetAllTasksinBucket), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetBucket), BucketCategory);
            builder.AddCustomAttributes(typeof(GetBucket), new DesignerAttribute(typeof(GetBucketDesigner)));
            builder.AddCustomAttributes(typeof(GetBucket), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteTask), TaskCategory);
            builder.AddCustomAttributes(typeof(DeleteTask), new DesignerAttribute(typeof(DeleteTaskDesigner)));
            builder.AddCustomAttributes(typeof(DeleteTask), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DeleteBucket), BucketCategory);
            builder.AddCustomAttributes(typeof(DeleteBucket), new DesignerAttribute(typeof(DeleteBucketDesigner)));
            builder.AddCustomAttributes(typeof(DeleteBucket), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CreateTask), TaskCategory);
            builder.AddCustomAttributes(typeof(CreateTask), new DesignerAttribute(typeof(CreateTaskDesigner)));
            builder.AddCustomAttributes(typeof(CreateTask), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(Createbucket), BucketCategory);
            builder.AddCustomAttributes(typeof(Createbucket), new DesignerAttribute(typeof(CreatebucketDesigner)));
            builder.AddCustomAttributes(typeof(Createbucket), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(UpdatePlan), PlanCategory);
            builder.AddCustomAttributes(typeof(UpdatePlan), new DesignerAttribute(typeof(UpdatePlanDesigner)));
            builder.AddCustomAttributes(typeof(UpdatePlan), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(UpdatePlanDetails), PlanCategory);
            builder.AddCustomAttributes(typeof(UpdatePlanDetails), new DesignerAttribute(typeof(UpdatePlanDetailsDesigner)));
            builder.AddCustomAttributes(typeof(UpdatePlanDetails), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(UpdateTask), TaskCategory);
            builder.AddCustomAttributes(typeof(UpdateTask), new DesignerAttribute(typeof(UpdateTaskDesigner)));
            builder.AddCustomAttributes(typeof(UpdateTask), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(UpdateTaskDetails), TaskCategory);
            builder.AddCustomAttributes(typeof(UpdateTaskDetails), new DesignerAttribute(typeof(UpdateTaskDetailsDesigner)));
            builder.AddCustomAttributes(typeof(UpdateTaskDetails), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(UpdateBucket), BucketCategory);
            builder.AddCustomAttributes(typeof(UpdateBucket), new DesignerAttribute(typeof(UpdateBucketDesigner)));
            builder.AddCustomAttributes(typeof(UpdateBucket), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(PostCommentToConversation), ConversationCategory);
            builder.AddCustomAttributes(typeof(PostCommentToConversation), new DesignerAttribute(typeof(PostCommentToConversationDesigner)));
            builder.AddCustomAttributes(typeof(PostCommentToConversation), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
