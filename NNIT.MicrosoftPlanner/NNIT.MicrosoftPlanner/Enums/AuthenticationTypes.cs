using System;
using System.ComponentModel;

namespace NNIT.MircrosoftPlanner.Enums
{
    public enum AuthenticationTypes
    {
        [Description("Integrated Windows Authentication")]
        IntegratedWindowsAuthentication,

        [Description("Credential Authentication")]
        UserAndPasswordAuthentication,

        [Description("Iteractive Authentication")]
        InteractiveAuthentication
    }
}
