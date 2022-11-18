using System;
using Microsoft.Xrm.Sdk;

namespace PluginTemplateWithExtensions.Helpers;

public class PluginContextHelper
{
    /// <summary>
    /// </summary>
    /// <param name="serviceProvider">passed as the argument to Execute method in the plugin</param>
    public PluginContextHelper(IServiceProvider serviceProvider)
    {
        ServiceProvider = serviceProvider;
        Context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));
        TracingService = (ITracingService) serviceProvider.GetService(typeof(ITracingService));
        ServiceFactory = (IOrganizationServiceFactory) serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        OrgService = ServiceFactory.CreateOrganizationService(null);
        UserService = ServiceFactory.CreateOrganizationService(Context.UserId);
    }

    public IServiceProvider ServiceProvider { get; }
    public IPluginExecutionContext Context { get; }
    public ITracingService TracingService { get; }
    public IOrganizationServiceFactory ServiceFactory { get; }

    /// <summary>
    ///     The IOrganizationService for the SYSTEM account which has read & write access to all records.
    ///     Use this service if need to read data and you want to avoid any issues with permissions.
    /// </summary>
    public IOrganizationService OrgService { get; }

    /// <summary>
    ///     The IOrganizationService for the user defined in the Plugin Step (could be Initiating user, SYSTEM, or designated
    ///     user).
    ///     Use this service if you writing data.
    /// </summary>
    public IOrganizationService UserService { get; }

    public Entity Target
    {
        get
        {
            if (!Context.InputParameters.Contains("Target"))
            {
                throw new Exception($"Plugin is trying to access the 'Target' InputParameter, but that doesn't exist in this messages [{Context.PrimaryEntityName}.{Context.MessageName}]");
            }

            return (Entity) Context.InputParameters["Target"];
        }
    }

    /// <summary>
    ///     Shorthand for Context.PostEntityImages["Target"].  also throws a nicer error if the the image wasn't registered
    /// </summary>
    public Entity PostImage
    {
        get
        {
            if (!Context.PostEntityImages.Contains("Target"))
            {
                throw new Exception($"Plugin is trying to access a Post Plugin Step Image named 'Target', but that doesn't exist in this messages [{Context.PrimaryEntityName}.{Context.MessageName}]");
            }

            return Context.PostEntityImages["Target"];
        }
    }

    /// <summary>
    ///     Returns a null if Context.PostEntityImages["Target"] doesn't exist
    /// </summary>
    public Entity PostImageOrDefault
    {
        get
        {
            if (!Context.PostEntityImages.Contains("Target"))
            {
                return null;
            }

            return Context.PostEntityImages["Target"];
        }
    }

    /// <summary>
    ///     Shorthand for Context.PreEntityImages["Target"].  also throws a nicer error if the the image wasn't registered
    /// </summary>
    public Entity PreImage
    {
        get
        {
            if (!Context.PreEntityImages.Contains("Target"))
            {
                throw new Exception($"Plugin is trying to access a Pre Plugin Step Image named 'Target', but that doesn't exist in this messages [{Context.PrimaryEntityName}.{Context.MessageName}]");
            }

            return Context.PreEntityImages["Target"];
        }
    }

    /// <summary>
    ///     Returns a null if Context.PreEntityImages["Target"] doesn't exist
    /// </summary>
    public Entity PreImageOrDefault
    {
        get
        {
            if (!Context.PreEntityImages.Contains("Target"))
            {
                return null;
            }

            return Context.PreEntityImages["Target"];
        }
    }


    public void Trace(string msg)
    {
        TracingService.Trace(msg);
    }


    #region plugin MessageNames

    public bool IsAssociate => Context.MessageName.ToLower() == "associate";

    public bool IsUpdate => Context.MessageName.ToLower() == "update";

    public bool IsDisassociate => Context.MessageName.ToLower() == "disassociate";

    public bool IsCreate => Context.MessageName.ToLower() == "create";
    public bool IsWin => Context.MessageName.ToLower() == "win";


    public bool IsPreValidation => Context.Stage == 10;
    public bool IsPreOperation => Context.Stage == 20;

    public bool IsPostOperation => Context.Stage == 40;

    public object StageName
    {
        get
        {
            switch (Context.Stage)
            {
                case 10:
                    return "Pre-Validation";
                case 20:
                    return "Pre-Operation";
                case 40:
                    return "Post-Operation";
                default:
                    return Context.Stage.ToString();
            }
        }
    }

    #endregion

    public void ValidateIsCreate()
    {
        if (!this.IsCreate)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered for the Create message, currently running on [{this.Context.MessageName}]");
        }
    }
    public void ValidateIsUpdate()
    {
        if (!this.IsUpdate)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered for the Update message, currently running on [{this.Context.MessageName}]");
        }
    }
    public void ValidateIsPreValidation()
    {
        if(!this.IsPreValidation)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered for the Update message, currently running on [{this.Context.MessageName}]");
        }
    }
    public void ValidateIsCreateOrUpdate()
    {
        if (!this.IsCreate && !this.IsUpdate)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered for the Create & Update message, currently running on [{this.Context.MessageName}]");
        }
    }
    public void ValidateIsWin()
    {
        if (!this.IsWin)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered for the Win message, currently running on [{this.Context.MessageName}]");
        }
    }

    public void ValidatePrimaryEntity<T>() where T : Entity
    {
        var name = typeof(T).Name;
        if (this.Context.PrimaryEntityName != typeof(T).Name.ToLower())
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered on the {name} entity, currently executing on the [{this.Context.PrimaryEntityName}]");
        }
    }
    public void ValidateIsPreOperation()
    {
        if (!this.IsPreOperation)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered on the Pre-Operation Stage, currently executing on [{this.StageName}] stage");
        }
    }
    public void ValidateIsPostOpertaion()
    {
        if (!this.IsPostOperation)
        {
            throw new InvalidPluginExecutionException($"This plugin should only be registered on the Post-Operation Stage, currently executing on [{this.StageName}] stage");
        }
    }
}