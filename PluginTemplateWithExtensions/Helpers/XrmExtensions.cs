using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Xml.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace PluginTemplateWithExtensions.Helpers;

public static class XrmExtensions
{

    #region Entity Extensions
    public static T GetEntityOrDefault<T>(this IOrganizationService service, EntityReference entRef, params string[] columns) where T : Entity, new()
    {
        return entRef != null ? service.GetEntity<T>(entRef, new ColumnSet(columns)) : null;
    }
    public static T GetEntity<T>(this IOrganizationService service, EntityReference entRef, ColumnSet cols = null) where T : Entity
    {
        var entity = Activator.CreateInstance<T>();
        if (entity.LogicalName != entRef.LogicalName)
        {
            var name = nameof(entRef);
            throw new ArgumentException($"'{name}' EntityRef for '{entRef.LogicalName}' doesn't match generic<Type> of '{entity.LogicalName}'", name);
        }
        return service.GetEntity<T>(entRef.Id, cols);
    }
    public static T GetEntity<T>(this IOrganizationService service, Guid id, ColumnSet cols = null) where T : Entity
    {
        if (cols == null)
        {
            cols = new ColumnSet(true);
        }
        var ttype = typeof(T);
        var entityNameConst = ttype.GetField("EntityLogicalName");
        if (entityNameConst != null)
        {
            var entityName = (string)entityNameConst.GetValue(null);
            T entity = (T)service.Retrieve(entityName, id, cols);
            return entity;
        }
        return null;
    }

    public static T GetEntity<T>(this IOrganizationService service, EntityReference entRef, params string[] columns) where T : Entity
    {
        return service.GetEntity<T>(entRef, new ColumnSet(columns));
    }
    public static T GetEntity<T>(this IOrganizationService service, Guid id, params string[] columns) where T : Entity
    {
        return service.GetEntity<T>(id, new ColumnSet(columns));
    }

    public static Entity Clone(this Entity entity, params string[] fieldsToSkip)
    {
        var skip = fieldsToSkip.ToList();
        var clone = new Entity(entity.LogicalName);
        foreach (var item in entity.Attributes.Where(x => !skip.Contains(x.Key)))
        {
            clone.Attributes.Add(item.Key, item.Value);
        }
        return clone;
    }
    public static T Clone<T>(this Entity entity, params string[] fieldsToSkip) where T : Entity
    {
        var skip = fieldsToSkip.ToList();
        var clone = new Entity(entity.LogicalName);
        foreach (var item in entity.Attributes.Where(x => !skip.Contains(x.Key) && x.Value is not AliasedValue))
        {
            clone.Attributes.Add(item.Key, item.Value);
        }
        return clone.ToEntity<T>();
    }


    /// <summary>
    /// Creates entity record that is has only the updated fields.   Useful when you want to update a record but don't want to 'touch' the fields that aren't changing.
    /// </summary>
    /// <param name="existing">The existing entity as read from disk.</param>
    /// <param name="newRecord">Entity record that just contains the all the values you want to update</param>
    /// <returns></returns>
    public static Entity GetUpdateRecord(this Entity existing, Entity newRecord)
    {
        var target = new Entity(existing.LogicalName, existing.Id);
        foreach (var attribute in newRecord.Attributes)
        {
            target.Update(existing, attribute.Key, attribute.Value);
        }
        return target;
    }

    /// <summary>
    /// Creates entity record that is has only the updated fields.   Useful when you want to update a record but don't want to 'touch' the fields that aren't changing.
    /// </summary>
    /// <param name="existing">The existing entity as read from disk.</param>
    /// <param name="newRecord">Entity record that just contains the all the values you want to update</param>
    /// <returns></returns>
    public static T GetUpdateRecord<T>(this T existing, T newRecord) where T : Entity
    {
        return existing.GetUpdateRecord((Entity)newRecord).ToEntity<T>();
    }

    /// <summary>
    /// This method is typically used from a plugin where you have your (target) and a composite (combo of target and ondisk values).
    /// You would then use the method to only update the target & composite if the newValue is different.
    /// typically used  to limit the fields from showing up in the audit log and/or triggering workflows/flows/plugins 
    /// </summary>
    /// <param name="target"></param>
    /// <param name="composite"></param>
    /// <param name="fieldName"></param>
    /// <param name="value"></param>
    public static void Update(this Entity target, Entity composite, string fieldName, object value)
    {
        var newValue = value;
        if (newValue is Enum)
        {
            newValue = new OptionSetValue((int)value);
        }

        if (IsValueDifferent(composite, fieldName, newValue))
        {
            composite[fieldName] = newValue;
            target[fieldName] = newValue;
        }
    }

    public static bool IsValueDifferent(this Entity entity, string fieldName, object value)
    {
        object testValue = value;
        if (testValue is Enum)
        {
            testValue = new OptionSetValue(value: ((int)value));
        }
        var hasValue = entity.Attributes.ContainsKey(fieldName);
        return ((!hasValue && testValue != null)
                || (hasValue && entity[fieldName] == null && testValue != null)
                || (hasValue && entity[fieldName] != null && !entity[fieldName].Equals(testValue)));
    }
    /// <summary>
    /// Typically used to compare a preimage & target within an update plugin to see if the field is being changed.
    /// </summary>
    public static bool HasValueChanged(this Entity entity, Entity updatedEntity, string fieldName)
    {
        var results = false;
        if (updatedEntity.Attributes.ContainsKey(fieldName))
        {
            var value = updatedEntity.Attributes[fieldName];
            results = entity.IsValueDifferent(fieldName, value);
        }
        return results;
    }

    public static EntityReference ToEntityRefWithName(this Entity entity)
    {
        EntityReference result = null;
        if (entity != null)
        {
            result = entity.ToEntityReference();
            result.Name = entity.GetName();
        }
        return result;
    }

    public static Dictionary<Guid, Entity> Add(this Dictionary<Guid, Entity> dictionary, Entity entity)
    {
        dictionary.Add(entity.Id, entity);
        return dictionary;
    }

    private static readonly List<string> _possibleNames = new List<string>() { "name", "subject", "title", "fullname", "lastname" };
    public static string GetName(this Entity entity)
    {
        string results = String.Empty;
        if (entity != null)
        {
            var nameField = _possibleNames.FirstOrDefault(entity.Contains);
            if (nameField == null && entity.LogicalName != null)
            {
                // Check for custom entity
                var parts = entity.LogicalName.Split('_');
                if (parts.Length > 1)
                {
                    var name = parts[0] + "_name";
                    if (entity.Contains(name))
                    {
                        nameField = name;
                    }
                }
            }

            if (nameField != null && entity[nameField] != null)
            {
                results = entity[nameField].ToString();
            }
        }
        return results;
    }

    public static bool ContainsAttributeNamed(this Entity entity, params string[] attributeNames)
    {
        return attributeNames.Any(entity.Contains);
    }
    #endregion

    public static void AddCreateRequest(this ICollection<OrganizationRequest> list, Entity target)
    {
        list.Add(new CreateRequest { Target = target });
    }

    public static void AddUpdateRequest(this ICollection<OrganizationRequest> list, Entity target)
    {
        list.Add(new UpdateRequest { Target = target });
    }

    /// <summary>
    /// Environment variables will have a default value shared for all environments then a current value which can be specified in each environment. This returns the
    /// the current value if it exists and the default value if it does not. This value is stored as a string. 
    /// </summary>
    /// <param name="service"></param>
    /// <param name="schemaName"></param>
    // TODO add Early Bound classes
    //public static string GetEnvironmentVariableCurrentOrDefault(this IOrganizationService service, string schemaName)
    //{
    //    var context = new CrmServiceContext(service);
    //    var query =
    //        from definition in context.EnvironmentVariableDefinitionSet
    //        join variable in context.EnvironmentVariableValueSet on definition.EnvironmentVariableDefinitionId equals variable.EnvironmentVariableDefinitionId.Id into environmentVariableInfo
    //        from subVariable in environmentVariableInfo.DefaultIfEmpty()
    //        where definition.SchemaName == schemaName
    //        select new
    //        {
    //            CurrentValue = subVariable.Value,
    //            definition.DefaultValue
    //        };

    //    var result = query.Single();
    //    var variableValue = result.CurrentValue ?? result.DefaultValue;

    //    return variableValue;
    //}


    private static object _optionSetLock = new object();
    private static Dictionary<string, EnumAttributeMetadata> _optionSetCache = new Dictionary<string, EnumAttributeMetadata>();
    public static string GetOptionSetLabel(this IOrganizationService service, string entityName, string fieldName, OptionSetValue optionSetValue)
    {
        if (optionSetValue == null)
            return "";

        var key = entityName + "." + fieldName;
        if (!_optionSetCache.ContainsKey(key))
        {
            lock (_optionSetLock)
            {
                if (!_optionSetCache.ContainsKey(key))
                {
                    var attReq = new RetrieveAttributeRequest();
                    attReq.EntityLogicalName = entityName;
                    attReq.LogicalName = fieldName;
                    attReq.RetrieveAsIfPublished = true;

                    var attResponse = (RetrieveAttributeResponse)service.Execute(attReq);
                    _optionSetCache.Add(key, (EnumAttributeMetadata)attResponse.AttributeMetadata);
                }
            }
        }
        var attMetadata = _optionSetCache[key];
        return attMetadata.OptionSet.Options.Where(x => x.Value == optionSetValue.Value).FirstOrDefault()?.Label.UserLocalizedLabel.Label;
    }


    public static string ToStringForLogging(this object value)
    {
        var result = string.Empty;
        try
        {
            if (value == null)
            {
                result = "NULL";
            }
            else if (value is EntityReference entRef)
            {
                if (entRef.Name != null)
                {
                    result += $"'{entRef.Name}'";
                }
                result += $"({entRef.LogicalName} {entRef.Id})";
            }
            else if (value is Money money)
            {
                result = money.Value.ToString(CultureInfo.InvariantCulture);
            }
            else if (value is OptionSetValue optionSet)
            {
                result = optionSet.Value.ToString();
            }
            else if (value is EntityCollection collection)
            {
                foreach (var entity in collection.Entities)
                {
                    result += $"({entity.LogicalName} {entity.Id} {entity.GetName()})";
                }
            }
            else if (value is Entity entity)
            {
                var name = entity.GetName();
                if (name != null)
                {
                    result += $"'{name}'";
                }
                result += $" ({entity.LogicalName} {entity.Id})";
            }
            else
            {
                result = value.ToString();
            }
        }
        catch
        {
            result = "[Unable to format value]";
        }
        return result;
    }

    public static bool HasRole(this IOrganizationService service, Guid userId, params Guid[] roleIds)
    {
        var values = new StringBuilder();
        foreach (var roleId in roleIds)
        {
            values.Append($"<value>{roleId}</value>");
        }


        var fetch = $@"
<fetch aggregate='true' >
  <entity name='systemuserroles' >
    <attribute name='systemuserroleid' alias='count' aggregate='count' />
    <filter>
      <condition attribute='systemuserid' operator='eq' value='{userId}' />
    </filter>
    <link-entity name='role' from='roleid' to='roleid' alias='r' >
      <filter>
        <condition attribute='parentrootroleid' operator='in' >
            {values}
        </condition>
      </filter>
    </link-entity>
  </entity>
</fetch>
";

        var query = new FetchExpression(fetch);
        var results = service.RetrieveMultiple(query);
        var count = results.Entities[0].GetAttributeValue<AliasedValue>("count");
        return count?.Value != null && (int)(count.Value) > 0;
    }

    #region DateTime Extensions

    //public static int? GetUsersTimeZoneSettings(this IOrganizationService service, Guid? userId = null)
    //{
    //    int? result = null;
    //    var condition = userId.HasValue
    //        ? new ConditionExpression(SystemUser.Fields.SystemUserId, ConditionOperator.Equal, userId)
    //        : new ConditionExpression(SystemUser.Fields.SystemUserId, ConditionOperator.EqualUserId);

    //    var response = service.RetrieveMultiple(
    //        new QueryExpression(UserSettings.EntityLogicalName)
    //        {
    //            ColumnSet = new ColumnSet(UserSettings.Fields.LocaleId, UserSettings.Fields.TimeZoneCode),
    //            Criteria = new FilterExpression
    //            {
    //                Conditions = { condition }
    //            }
    //        });
    //    if (response.Entities is {Count: > 0})
    //    {
    //        var timeZoneCode = response.Entities[0].ToEntity<UserSettings>().TimeZoneCode;
    //        result = timeZoneCode;
    //    }
    //    return result;
    //}
    public static DateTime RetrieveLocalTimeFromUtcTime(this IOrganizationService service, DateTime utcTime, int timeZoneCode)
    {
        if (utcTime.Kind == DateTimeKind.Unspecified)
        {
            throw new InvalidEnumArgumentException($"{nameof(RetrieveLocalTimeFromUtcTime)} only supports UTC times, the provided DateTime has a kind of {utcTime.Kind}");
        }
        if (utcTime.Kind == DateTimeKind.Local)
        {
            utcTime = utcTime.ToUniversalTime();
        }
        var request = new LocalTimeFromUtcTimeRequest
        {
            TimeZoneCode = timeZoneCode,
            UtcTime = utcTime
        };
        var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
        return response.LocalTime;
    }
    public static DateTime RetrieveUtcTimeFromLocalTime(this IOrganizationService service, DateTime localTime, int timeZoneCode)
    {
        if (localTime.Kind == DateTimeKind.Unspecified)
        {
            throw new InvalidEnumArgumentException($"{nameof(RetrieveUtcTimeFromLocalTime)} only supports LOCAL times, the provided DateTime has a kind of {localTime.Kind}");
        }
        if (localTime.Kind == DateTimeKind.Utc)
        {
            localTime = localTime.ToLocalTime();
        }
        var request = new UtcTimeFromLocalTimeRequest()
        {
            TimeZoneCode = timeZoneCode,
            LocalTime = localTime
        };
        var response = (UtcTimeFromLocalTimeResponse)service.Execute(request);
        return response.UtcTime;
    }

    //public static DateTime GetUsersCurrentTime(this IOrganizationService service, Guid? userId = null)
    //{
    //    return service.GetUsersTime(DateTime.UtcNow, userId);
    //}
    //public static DateTime GetUsersTime(this IOrganizationService service, DateTime utcTime, Guid? userId = null)
    //{
    //    var timeZoneCode = service.GetUsersTimeZoneSettings(userId);
    //    return ConvertToLocalTime(service, utcTime, timeZoneCode);
    //}
    public static DateTime ConvertToLocalTime(this IOrganizationService service, DateTime utcTime, int? timeZoneCode)
    {
        return timeZoneCode.HasValue
            ? service.RetrieveLocalTimeFromUtcTime(utcTime, timeZoneCode.Value)
            : utcTime;
    }
    public static DateTime? ConvertToLocalTime(this IOrganizationService service, DateTime? utcTime, int? timeZoneCode)
    {
        return timeZoneCode.HasValue && utcTime.HasValue
            ? service.RetrieveLocalTimeFromUtcTime(utcTime.Value, timeZoneCode.Value)
            : utcTime;
    }
    public static DateTime ConvertToUtcTime(this IOrganizationService service, DateTime localTime, int? timeZoneCode)
    {
        return timeZoneCode.HasValue
            ? service.RetrieveUtcTimeFromLocalTime(localTime, timeZoneCode.Value)
            : localTime;
    }
    #endregion

    public static string GetLogicalNameOfT<T>() where T : Entity
    {
        var ttype = typeof(T);
        var entityNameConst = ttype.GetField("EntityLogicalName");
        return (string)entityNameConst.GetValue(null);
    }
    public static void IsA<T>(this EntityReference entityReference, string argumentName = null) where T : Entity
    {
        if (entityReference != null)
        {
            var entityName = GetLogicalNameOfT<T>();
            if (entityReference.LogicalName != entityName)
            {
                throw new InvalidEnumArgumentException($"{argumentName} is not a {entityName}");
            }
        }
    }

    /// <summary>
    /// ExecuteMultipleRequest is limited to 1000 requests,  this extension method will allow to execute more then 1000 requests,
    /// by breaking them up into batches and submitting each batch one at a time.
    ///
    /// This will still be limited to the 2 min plugin/custom workflow activity timeout.
    /// 
    /// If any error is detected, an error will be thrown.
    ///
    /// https://docs.microsoft.com/en-us/powerapps/developer/data-platform/org-service/execute-multiple-requests#run-time-limitations
    /// </summary>
    /// <param name="service"></param>
    /// <param name="trace"></param>
    /// <param name="requests"></param>
    /// <param name="batchSize">default is 1000 and should only be set to a smaller number</param>
    /// <returns></returns>
    public static OrganizationResponse[] ExecuteMultiple(this IOrganizationService service, ITracingService trace, OrganizationRequest[] requests, int batchSize = 1000)
    {
        return ExecuteBatches(trace, requests, batchSize,
            (batch) =>
            {
                var requestCollection = new OrganizationRequestCollection();
                requestCollection.AddRange(batch);
                var request = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    Requests = requestCollection,
                };
                return request;
            },
            (request) =>
            {
                var executeMultipleResponse = (ExecuteMultipleResponse)service.Execute(request);
                if (executeMultipleResponse.IsFaulted)
                {
                    var firstError = executeMultipleResponse.Responses.First(responseItem => responseItem.Fault != null);
                    throw new InvalidPluginExecutionException($"ExecuteMultipleRequest failed; {firstError.Fault.Message}");
                }
                return executeMultipleResponse.Responses.OrderBy(response => response.RequestIndex).Select(response => response.Response).ToArray();
            });
    }


    /// <summary>
    /// ExecuteTransactionRequest is limited to 1000 requests,  this extension method will allow to execute more then 1000 requests,
    /// by breaking them up into batches and submitting each batch one at a time.
    /// </summary>
    /// <param name="service"></param>
    /// <param name="trace"></param>
    /// <param name="requests"></param>
    /// <param name="batchSize">default is 1000 and should only be set to a smaller number for testing</param>
    /// <returns></returns>
    public static OrganizationResponse[] ExecuteTransaction(this IOrganizationService service, ITracingService trace, OrganizationRequest[] requests, int batchSize = 1000)
    {
        return ExecuteBatches(trace, requests, batchSize,
            (batch) =>
            {
                var requestCollection = new OrganizationRequestCollection();
                requestCollection.AddRange(batch);
                var request = new ExecuteTransactionRequest()
                {
                    ReturnResponses = true,
                    Requests = requestCollection,
                };
                return request;
            },
            (request) =>
            {
                var response = (ExecuteTransactionResponse)service.Execute(request);
                return response.Responses.ToArray();
            });
    }

    private static OrganizationResponse[] ExecuteBatches(
        ITracingService trace,
        OrganizationRequest[] requests,
        int batchSize,
        Func<OrganizationRequest[], OrganizationRequest> convertBatchToRequest, Func<OrganizationRequest, OrganizationResponse[]> execute)
    {
        var results = new List<OrganizationResponse>();
        var count = requests.Length;
        if (count > 0)
        {
            var totalBatches = Math.Ceiling((count + 0m) / (batchSize + 0m));
            trace.Trace($"Executing {count} requests in {totalBatches} batch(es)");
            for (var i = 0; i < totalBatches; i++)
            {
                var offset = i * batchSize;
                var batch = requests.Skip(offset).Take(batchSize).ToArray();
                var request = convertBatchToRequest(batch);
                trace.Trace($"Executing batch {i + 1} of {totalBatches}");

                var responses = execute(request);
                results.AddRange(responses);
            }
        }
        return results.ToArray();
    }

    public static Guid[] GetIdsFromCreatedResponses(this OrganizationResponse[] responses)
    {
        var results = responses.Where(response => response is CreateResponse)
            .Select(response => ((CreateResponse)response).id);
        return results.ToArray();
    }

    public static Guid[] GetIds(this OrganizationResponse[] responses, OrganizationRequest[] requests)
    {
        var results = responses.Select((response, index) =>
        {
            Guid id;
            switch (response)
            {
                case CreateResponse createResponse:
                    id = createResponse.id;
                    break;
                case UpdateResponse _:
                    var originalRequest = (UpdateRequest)requests[index];
                    id = originalRequest.Target.Id;
                    break;
                default:
                    throw new InvalidPluginExecutionException($"Did not expect a results of type {response.GetType().Name}");
            }
            return id;
        });
        return results.ToArray();
    }
    /// <summary>
    /// Layers target on top of original to give you a Composite of what entity will look like after the save has completed.  Typically used in pre-operation or pre-validation of a plugin
    /// </summary>
    /// <typeparam name="TEntity">CRM early bound entity class eg account, contact, opportunity</typeparam>
    /// <param name="original"></param>
    /// <param name="target"></param>
    /// <returns></returns>
    public static TEntity BuildCompositeEntity<TEntity>(this TEntity target, TEntity original) where TEntity : Entity, new()
    {
        var composite = new TEntity
        {
            Id = original.Id
        };
        foreach (var attribute in original.Attributes) composite[attribute.Key] = attribute.Value;
        foreach (var attribute in target.Attributes) composite[attribute.Key] = attribute.Value;
        return composite;
    }

    /// <summary>
    /// Allows you to perform an 'In' clause when using the dynamics linq provider
    /// 
    /// https://stackoverflow.com/a/57715528/6522
    /// </summary>
    public static IQueryable<TSource> WhereIn<TSource, TValue>(this IQueryable<TSource> source, Expression<Func<TSource, TValue>> valueSelector, IEnumerable<TValue> values)
    {
        if (null == source) { throw new ArgumentNullException(nameof(source)); }
        if (null == valueSelector) { throw new ArgumentNullException(nameof(valueSelector)); }
        if (null == values) { throw new ArgumentNullException(nameof(values)); }

        var equalExpressions = new List<BinaryExpression>();

        foreach (var value in values)
        {
            var equalsExpression = Expression.Equal(valueSelector.Body, Expression.Constant(value));
            equalExpressions.Add(equalsExpression);
        }

        var parameter = valueSelector.Parameters.Single();
        var combined = equalExpressions.Aggregate<Expression>((accumulate, equal) => Expression.Or(accumulate, equal));
        var combinedLambda = Expression.Lambda<Func<TSource, bool>>(combined, parameter);

        return source.Where(combinedLambda);
    }

    public static double? ToDouble(this decimal? value)
    {
        return value.HasValue ? (double)value.Value : default;
    }

    /// <summary>
    /// Prints the FetchXML request of a query to the trace log. To use when testing in CE.
    /// </summary>
    public static void TraceFetchXmlForTesting(this IQueryable query, IOrganizationService orgService, ITracingService tracingService)
    {
        var queryProvider = query.Provider;
        var translateMethodInfo = queryProvider.GetType().GetMethod("Translate");
        var queryEx = (QueryExpression)translateMethodInfo.Invoke(queryProvider, new object[] { query.Expression });

        var reqConvertToFetchXml = new QueryExpressionToFetchXmlRequest { Query = queryEx };
        var respConvertToFetchXml = (QueryExpressionToFetchXmlResponse)orgService.Execute(reqConvertToFetchXml);
        var xml = respConvertToFetchXml.FetchXml;
        var formattedXml = XDocument.Parse(xml).ToString();

        tracingService.Trace("To FetchXML:" + Environment.NewLine);
        tracingService.Trace(formattedXml);
    }
}