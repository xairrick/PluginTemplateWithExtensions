using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PluginTemplateWithExtensions.Helpers;

public static class PluginExtensions
{
    public const string OUTPUT_PARAMETER_BUSINESS_ENTITY = "BusinessEntity";
    public const string OUTPUT_PARAMETER_BUSINESS_ENTITY_COLLECTION = "BusinessEntityCollection";

    public const string INPUT_PARAMETERS_QUERY = "Query";
    public const string INPUT_PARAMETERS_COLUMNSET = "ColumnSet";

    public static IPluginExecutionContext Context(this IServiceProvider serviceProvider)
    {
        return (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
    }

    public static ITracingService TracingService(this IServiceProvider serviceProvider)
    {
        return (ITracingService)serviceProvider.GetService(typeof(ITracingService));
    }

    public static IOrganizationServiceFactory ServiceFactory(this IServiceProvider serviceProvider)
    {
        return (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
    }

    /// <summary>
    ///     The IOrganizationService for the SYSTEM account which has read & write access to all records.
    ///     Use this service if need to read data and you want to avoid any issues with permissions.
    /// </summary>
    public static IOrganizationService OrgService(this IServiceProvider serviceProvider)
    {
        return serviceProvider.ServiceFactory().CreateOrganizationService(null);
    }

    /// <summary>
    ///     The IOrganizationService for the user defined in the Plugin Step (could be Initiating user, SYSTEM, or designated
    ///     user).
    ///     Use this service if you writing data.
    /// </summary>
    public static IOrganizationService UserService(this IServiceProvider serviceProvider)
    {
        return serviceProvider.ServiceFactory().CreateOrganizationService(serviceProvider.Context().UserId);
    }

    public static TEntity Target<TEntity>(this IPluginExecutionContext context) where TEntity : Entity =>
        !context.InputParameters.Contains("Target")
            ? throw new InvalidPluginExecutionException(
                $"Plugin is trying to access the 'Target' InputParameter, but that doesn't exist in this messages [{context.PrimaryEntityName}.{context.MessageName}]")
            : ((Entity)context.InputParameters["Target"]).ToEntity<TEntity>();

    /// <summary>
    ///     Shorthand for Context.PostEntityImages["Target"].  also throws a nicer error if the the image wasn't registered
    /// </summary>
    public static TEntity PostImage<TEntity>(this IPluginExecutionContext context) where TEntity : Entity =>
        !context.PostEntityImages.Contains("Target")
            ? throw new InvalidPluginExecutionException(
                $"Plugin is trying to access a Post Plugin Step Image named 'Target', but that doesn't exist in this messages [{context.PrimaryEntityName}.{context.MessageName}]")
            : context.PostEntityImages["Target"].ToEntity<TEntity>();

    /// <summary>
    ///     Returns a null if Context.PostEntityImages["Target"] doesn't exist
    /// </summary>
    public static TEntity PostImageOrDefault<TEntity>(this IPluginExecutionContext context) where TEntity : Entity =>
        !context.PostEntityImages.Contains("Target")
            ? null
            : context.PostEntityImages["Target"].ToEntity<TEntity>();

    /// <summary>
    ///     Shorthand for Context.PreEntityImages["Target"].  also throws a nicer error if the the image wasn't registered
    /// </summary>
    public static TEntity PreImage<TEntity>(this IPluginExecutionContext context) where TEntity : Entity =>
        !context.PreEntityImages.Contains("Target")
            ? throw new InvalidPluginExecutionException(
                $"Plugin is trying to access a Pre Plugin Step Image named 'Target', but that doesn't exist in this messages [{context.PrimaryEntityName}.{context.MessageName}]")
            : context.PreEntityImages["Target"].ToEntity<TEntity>();

    /// <summary>
    ///     Returns a null if Context.PreEntityImages["Target"] doesn't exist
    /// </summary>
    public static TEntity PreImageOrDefault<TEntity>(this IPluginExecutionContext context) where TEntity : Entity =>
        !context.PreEntityImages.Contains("Target")
            ? null
            : context.PreEntityImages["Target"].ToEntity<TEntity>();

    public static Entity RetrievedEntity(this IPluginExecutionContext context)
    {
        return (Entity)context.OutputParameters[OUTPUT_PARAMETER_BUSINESS_ENTITY];
    }

    public static TEntity RetrievedEntity<TEntity>(this IPluginExecutionContext context)
        where TEntity : Entity
    {
        return context.RetrievedEntity().ToEntity<TEntity>();
    }

    public static EntityCollection RetrievedEntities(this IPluginExecutionContext context)
    {
        return (EntityCollection)context.OutputParameters[OUTPUT_PARAMETER_BUSINESS_ENTITY_COLLECTION];
    }
    public static void SetRetrievedEntities(this IPluginExecutionContext context, EntityCollection collection)
    {
        context.OutputParameters[OUTPUT_PARAMETER_BUSINESS_ENTITY_COLLECTION] = collection;
    }

    public static QueryBase Query(this IPluginExecutionContext context)
    {
        return (QueryBase)context.InputParameters[INPUT_PARAMETERS_QUERY];
    }
    public static ColumnSet ColumnSet(this IPluginExecutionContext context)
    {
        return (ColumnSet)context.InputParameters[INPUT_PARAMETERS_COLUMNSET];
    }

    public static void SetQuery(this IPluginExecutionContext context, QueryBase query)
    {
        context.InputParameters[INPUT_PARAMETERS_QUERY] = query;
    }


    public static void Trace(this ITracingService tracingService, string msg)
    {
        tracingService.Trace(msg);
    }

    private static void _ValidateMessage(this IPluginExecutionContext context, bool condition, string expectedMessage)
    {
        if (!condition)
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered for the {expectedMessage} message, currently running on [{context.MessageName}]");
        }
    }

    public static void ValidateIsCreate(this IPluginExecutionContext context)
    {
        context._ValidateMessage(context.IsCreate(), "Create");
    }

    public static void ValidateIsUpdate(this IPluginExecutionContext context)
    {
        context._ValidateMessage(context.IsUpdate(), "Update");
    }

    public static void ValidateIsPreValidation(this IPluginExecutionContext context)
    {
        context._ValidateMessage(context.IsPreValidation(), "PreValidation");
    }

    public static void ValidateIsCreateOrIsUpdate(this IPluginExecutionContext context)
    {
        context._ValidateMessage((context.IsCreate() || context.IsUpdate()), "Create or Update");
    }

    public static void ValidateIsWin(this IPluginExecutionContext context)
    {
        if (!context.IsWin())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered for the Win message, currently running on [{context.MessageName}]  ");
        }
    }

    public static void ValidatePrimaryEntity<TEntity>(this IPluginExecutionContext context) where TEntity : Entity
    {
        var name = typeof(TEntity).Name;
        if (context.PrimaryEntityName != typeof(TEntity).Name.ToLower())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered on the {name} entity, currently executing on the [{context.PrimaryEntityName}]");
        }
    }

    public static void ValidateIsPreOperation(this IPluginExecutionContext context)
    {
        if (!context.IsPreOperation())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered on the Pre-Operation Stage, currently executing on [{context.StageName()}] stage");
        }
    }

    public static void ValidateIsPostOperation(this IPluginExecutionContext context)
    {
        if (!context.IsPostOperation())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered on the Post-Operation Stage, currently executing on [{context.StageName()}] stage");
        }
    }

    public static void ValidatePreImage(this IPluginExecutionContext context)
    {
        if (context.IsUpdate() && !context.PreEntityImages.ContainsKey("Target"))
        {
            throw new InvalidPluginExecutionException("Could not find preImage with name 'Target' in plugin step");
        }
    }

    public static void ValidateIsAsync(this IPluginExecutionContext context)
    {
        if (!context.IsRunningAsynchronous())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered for the Asynchronous Mode, currently executing in Synchronous mode ");
        }
    }

    public static bool IsAssociate(this IPluginExecutionContext context) => context.MessageName.ToLower() == "associate";

    public static bool IsUpdate(this IPluginExecutionContext context) => context.MessageName.ToLower() == "update";

    public static bool IsDisassociate(this IPluginExecutionContext context) => context.MessageName.ToLower() == "disassociate";

    public static bool IsCreate(this IPluginExecutionContext context) => context.MessageName.ToLower() == "create";

    public static bool IsRetrieve(this IPluginExecutionContext context) => context.MessageName.ToLower() == "retrieve";
    public static bool IsRetrieveMultiple(this IPluginExecutionContext context) => context.MessageName.ToLower() == "retrievemultiple";
    public static bool IsWin(this IPluginExecutionContext context) => context.MessageName.ToLower() == "win";


    public static bool IsPreValidation(this IPluginExecutionContext context) => context.Stage == 10;
    public static bool IsPreOperation(this IPluginExecutionContext context) => context.Stage == 20;

    public static bool IsPostOperation(this IPluginExecutionContext context) => context.Stage == 40;
    public static bool IsRunningAsynchronous(this IPluginExecutionContext context) => context.Mode == 1;

    public static string StageName(this IPluginExecutionContext context)
    {
        switch (context.Stage)
        {
            case 10:
                return "Pre-Validation";
            case 20:
                return "Pre-Operation";
            case 40:
                return "Post-Operation";
            default:
                return context.Stage.ToString();
        }
    }
        
    /// <summary>
    /// Uses Eastern Time Zone to determine the date
    /// </summary>
    /// <param name="context">IPluginExecutionContext</param>
    /// <param name="organizationService">Used to lookup the user's TimeZone</param>
    /// <returns>Date Only (time will 12:00am)</returns>
    public static DateTime OperationCreatedOnUsersTimeZone(this IPluginExecutionContext context, IOrganizationService organizationService)
    {
        return organizationService.GetUsersTime(context.OperationCreatedOnUtc());
    }
    /// <summary>
    /// OperationCreatedOn is known to be UTC time, but the plugin context only provided it as unspecified.
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static DateTime OperationCreatedOnUtc(this IPluginExecutionContext context)
    {
        //The OperationCreatedOn is known to be a UTC date with the Kind of Unspecified
        var operationCreatedOnUtc = DateTime.SpecifyKind(context.OperationCreatedOn, DateTimeKind.Utc);
        return operationCreatedOnUtc;
    }

    /// <summary>
    /// Returns the Date based on midnight of the given time zone<br/>
    /// If your are passing in an Eastern time zone, then 4am/5am GMT will be midnight cutoff<br/>
    /// eg  2022-09-21T03:00:00Z  will  return 2022-09-20  (as it's still the 20th in the eastern time zone)<br/>
    /// </summary>
    /// <param name="context">IPluginExecutionContext</param>
    /// <param name="timeZone"></param>
    /// <returns>Date Only (time will 12:00am)</returns>
    public static DateTime OperationCreatedOnDate(this IPluginExecutionContext context, TimeZoneInfo timeZone)
    {
        return TimeZoneInfo.ConvertTime(context.OperationCreatedOnUtc(), timeZone).Date;
    }
        
    public static void ValidateIsSync(this IPluginExecutionContext context)
    {
        if (context.IsRunningAsynchronous())
        {
            throw new InvalidPluginExecutionException(
                $"This plugin should only be registered for the Synchronous Mode, currently executing in Asynchronous mode ");
        }
    }
}