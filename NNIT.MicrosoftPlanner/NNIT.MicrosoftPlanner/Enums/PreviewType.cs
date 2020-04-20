using System;
using System.ComponentModel;

namespace NNIT.MircrosoftPlanner.Enums
{
    public enum PreviewTypes
    {
        [Description("")]
        NoChange,

        [Description("No Preview")]
        noPreview,

        [Description("Automatic")]
        automatic,

        [Description("Checklist")]
        checklist,

        [Description("Description")]
        description,

        [Description("Reference")]
        reference
    }
}