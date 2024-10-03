using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

class Program
{
    // TODO Enter your Dataverse environment's URL and logon info.
    static string url = "https://bolt.crm.dynamics.com";
    static string userName = "lcallanan@nixonpower.com";
    static string password = "";

    // This service connection string uses the info provided above.
    // The AppId and RedirectUri are provided for sample code testing.
    static string connectionString = $@"
   AuthType = OAuth;
   Url = {url};
   UserName = {userName};
   Password = {password};
   LoginPrompt=Auto;
   RequireNewInstance = True";

    static void Main()
    {
        //ServiceClient implements IOrganizationService interface
        IOrganizationService service = new ServiceClient(connectionString);

        FetchXmlExample(service);
    }


    /// <summary>
    /// Demonstrates RetrieveMultiple with FetchXml
    /// </summary>
    /// <param name="service">Authenticated IOrganizationService instance</param>
    static void FetchXmlExample(IOrganizationService service)
    {
        // Define the query
        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1000'>
  <entity name='asyncoperation'>
    <attribute name='asyncoperationid' />
    <attribute name='name' />
    <attribute name='regardingobjectid' />
    <attribute name='operationtype' />
    <attribute name='statuscode' />
    <attribute name='ownerid' />
    <attribute name='startedon' />
    <attribute name='statecode' />
    <attribute name='postponeuntil' />
    <attribute name='statuscode' />
    <attribute name='createdon' />
    <attribute name='createdby' />
    <order attribute='startedon' descending='true' />
    <filter type='and'>
      <condition attribute='statuscode' operator='in'>
        <value>20</value>
        <value>10</value>
        <value>0</value>
      </condition>
      <condition attribute='name' operator='eq' value='WF_ResidentialOwnerInvoicing' />
    </filter>
  </entity>
</fetch>";

        FetchExpression query = new FetchExpression(fetchXml);

        // Send the request
        EntityCollection results = service.RetrieveMultiple(query);

        // Create an ExecuteMultipleRequest object.
        var multipleRequest = new ExecuteMultipleRequest()
        {
            // Assign settings that define execution behavior: continue on error, return responses. 
            Settings = new ExecuteMultipleSettings()
            {
                ContinueOnError = true,
                ReturnResponses = false
            },
            // Create an empty organization request collection.
            Requests = new OrganizationRequestCollection()
        };

        // Show the data
        int index_value = 1;
        foreach (Entity record in results.Entities)
        {
            
            record["statuscode"] = new OptionSetValue(32);
            record["statecode"] = new OptionSetValue(3);
            try 
            {
                UpdateRequest updateRequest = new UpdateRequest { Target = record };
                multipleRequest.Requests.Add(updateRequest);

                //service.Update(record);
                Console.WriteLine(index_value);
                Console.WriteLine($"id:{record.GetAttributeValue<Guid>("asyncoperationid")}");
                Console.WriteLine($"status:{record.FormattedValues["statuscode"]}\n");
                index_value++;
            }
            catch 
            {
                Console.WriteLine(index_value);
                Console.WriteLine($"id:{record.GetAttributeValue<Guid>("asyncoperationid")}");
                Console.WriteLine($"status:{record.FormattedValues["statuscode"]}");
                Console.WriteLine($"erred\n");
                index_value++;
            }
        }

        try
        {
            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);
            Console.WriteLine($"Success\n");
        }
        catch { Console.WriteLine($"erred\n"); }
    }

    /// <summary>
    /// Call this method for bulk update
    /// </summary>
    /// <param name="service">Org Service</param>
    /// <param name="entities">Collection of entities to Update</param>
    public static void BulkUpdate(IOrganizationService service, DataCollection<Entity> entities)
    {
        // Create an ExecuteMultipleRequest object.
        var multipleRequest = new ExecuteMultipleRequest()
        {
            // Assign settings that define execution behavior: continue on error, return responses. 
            Settings = new ExecuteMultipleSettings()
            {
                ContinueOnError = false,
                ReturnResponses = true
            },
            // Create an empty organization request collection.
            Requests = new OrganizationRequestCollection()
        };

        // Add a UpdateRequest for each entity to the request collection.
        foreach (var entity in entities)
        {
            UpdateRequest updateRequest = new UpdateRequest { Target = entity };
            multipleRequest.Requests.Add(updateRequest);
        }

        // Execute all the requests in the request collection using a single web method call.
        ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);

    }
}