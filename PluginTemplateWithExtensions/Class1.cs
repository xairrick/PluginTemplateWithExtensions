using System;
using Microsoft.Xrm.Sdk;
using PluginTemplateWithExtensions.Helpers;

namespace PluginTemplateWithExtensions;

public class Class1 : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        var context = serviceProvider.Context();
        var trace = serviceProvider.TracingService();
        var orgService = serviceProvider.OrgService();
        var userService = serviceProvider.UserService();

        // TODO Early Bound Example -> context.ValidatePrimaryEntity<msdyn_agreement>();
        context.ValidateIsPreOperation();
        trace.Log("Confirmed that this plugins is running as Pre-Operation");

        // TODO Early Bound Example -> var preImage = context.PreImage<msdyn_agreement>();
        // TODO Early Bound Example -> var target = context.Target<msdyn_agreement>();

        var id = new Guid("b5b42f92-f9eb-4dc5-aec1-ec376d7137b9");
        // TODO Early Bound Example -> var account = orgService.GetEntity<Account>(id, "name", "description");

        var entityReference = new EntityReference("contact", new Guid("87702e69-92c7-4dd5-b89b-57e86b612557"));
        // TODO Early Bound Example -> var contact = orgService.GetEntity<Contact>(entityReference);  // Returns all fields
    }
}

