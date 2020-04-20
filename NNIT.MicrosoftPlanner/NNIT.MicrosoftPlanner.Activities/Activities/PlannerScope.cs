using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Security;
using System.Activities.Statements;
using System.ComponentModel;
using NNIT.MircrosoftPlanner.Enums;
using NNIT.MircrosoftPlanner.Activities;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Utilities;
using UiPath.Shared.Activities.Localization;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using NNIT.MicrosoftPlanner.Activities.Properties;

namespace NNIT.MircrosoftPlanner.Activities
{
    [LocalizedDisplayName(nameof(Resources.PlannerScope_DisplayName))]
    [LocalizedDescription(nameof(Resources.PlannerScope_Description))]
    public class PlannerScope : ContinuableAsyncNativeActivity
    {
        #region Properties

        [Browsable(false)]
        public ActivityAction<IObjectContainer​> Body { get; set; }

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

        [LocalizedDisplayName(nameof(Resources.PlannerScope_ClientId_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_ClientId_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        public InArgument<string> ClientId { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_TenantId_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_TenantId_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        public InArgument<string> TenantId { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_AuthenticationType_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_AuthenticationType_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        [TypeConverter(typeof(EnumNameConverter<AuthenticationTypes>))]
        public AuthenticationTypes AuthenticationType { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_Username_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_Username_Description))]
        [LocalizedCategory(nameof(Resources.Credential_Category))]
        public InArgument<string> Username { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_Password_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_Password_Description))]
        [LocalizedCategory(nameof(Resources.Credential_Category))]
        public InArgument<SecureString> Password { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_RedirectUri_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_RedirectUri_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        public InArgument<string> RedirectUri { get; set; }

        [LocalizedDisplayName(nameof(Resources.PlannerScope_AuthToken_DisplayName))]
        [LocalizedDescription(nameof(Resources.PlannerScope_AuthToken_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> AuthToken { get; set; }

        // A tag used to identify the scope in the activity context
        internal static string ParentContainerPropertyTag => "ScopeActivity";

        // Object Container: Add strongly-typed objects here and they will be available in the scope's child activities.
        private readonly IObjectContainer _objectContainer;

        private string _AccessToken;
        private string _clientId;
        private string _tenantId;
        private string _username;
        private SecureString _password;
        private string _redirecturi;
        #endregion


        #region Constructors

        public PlannerScope(IObjectContainer objectContainer)
        {
            _objectContainer = objectContainer;

            Body = new ActivityAction<IObjectContainer>
            {
                Argument = new DelegateInArgument<IObjectContainer>(ParentContainerPropertyTag),
                Handler = new Sequence { DisplayName = Resources.Do }
            };
        }

        public PlannerScope() : this(new ObjectContainer())
        {

        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            if (ClientId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ClientId)));
            if (TenantId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(TenantId)));
            if (AuthenticationType == AuthenticationTypes.UserAndPasswordAuthentication && (Username == null || Password == null ) ) metadata.AddValidationError(string.Format(Resources.ValidationAuthenticationMethod_Error, nameof(Username), nameof(Password)));
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<NativeActivityContext>> ExecuteAsync(NativeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var timeout = TimeoutMS.Get(context);
            _clientId = ClientId.Get(context);
            _tenantId = TenantId.Get(context);
            _username = Username.Get(context);
            _password = Password.Get(context);
            _redirecturi = RedirectUri.Get(context);

            // Set a timeout on the execution
            var task = ExecuteWithTimeout(context, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            return (ctx) => {
                // Schedule child activities
                if (Body != null)
                    ctx.ScheduleAction<IObjectContainer>(Body, _objectContainer, OnCompleted, OnFaulted);

                //send to Children
                _objectContainer.Add(_AccessToken);

                // Outputs
                AuthToken.Set(ctx, _AccessToken);
            };
        }

        private async Task ExecuteWithTimeout(NativeActivityContext context, CancellationToken cancellationToken = default)
        {

            string authority = string.Format("https://login.microsoftonline.com/{0}", _tenantId);
            AuthenticationResult result = (AuthenticationResult)null;
            AuthenticationContext ac = new AuthenticationContext(authority);

            switch (AuthenticationType)
            {
                case AuthenticationTypes.UserAndPasswordAuthentication:
                    string _username = Username.Get(context);
                    SecureString _password = Password.Get(context);
                    try
                    {
                        result = await ac.AcquireTokenAsync("https://graph.microsoft.com", _clientId, new UserPasswordCredential(_username, _password));
                        _AccessToken = result.AccessToken;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Something went wrong. Error message: " + ex.Message);
                    }
                    break;
                case AuthenticationTypes.InteractiveAuthentication:
                    try
                    {
                        result = await ac.AcquireTokenAsync("https://graph.microsoft.com", _clientId, new Uri(_redirecturi), new PlatformParameters(PromptBehavior.RefreshSession));
                        _AccessToken = result.AccessToken;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Something went wrong. Error message: " + ex.Message);
                    }
                    break;
                case AuthenticationTypes.IntegratedWindowsAuthentication:
                    try
                    {
                        result = await ac.AcquireTokenAsync("https://graph.microsoft.com", _clientId, new Uri(_redirecturi), new PlatformParameters(PromptBehavior.Auto));
                        _AccessToken = result.AccessToken;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Something went wrong. Error message: " + ex.Message);
                    }
                    break;
            }
        }

        #endregion


        #region Events

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            faultContext.CancelChildren();
            Cleanup();
        }

        private void OnCompleted(NativeActivityContext context, ActivityInstance completedInstance)
        {
            Cleanup();
        }

        #endregion


        #region Helpers

        private void Cleanup()
        {
            var disposableObjects = _objectContainer.Where(o => o is IDisposable);
            foreach (var obj in disposableObjects)
            {
                if (obj is IDisposable dispObject)
                    dispObject.Dispose();
            }
            _objectContainer.Clear();
        }

        #endregion
    }
}

