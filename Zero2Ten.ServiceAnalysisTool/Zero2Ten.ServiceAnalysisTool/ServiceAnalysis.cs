using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;

namespace Zero2Ten.ServiceAnalysis
{
    public class ServiceAnalysisPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracer.Trace("Start of Plugin Execution");

            Entity entity;

            // Check if the input parameters property bag contains a target
            // of the create operation and that target is of type Entity.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target business entity from the input parameters.
                entity = (Entity)context.InputParameters["Target"];
            }
            else
            {
                return;
            }

            try
            {
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                
                tracer.Trace("Start of Plugin Code");

                // TODO: Need the decimal values on Sales Analysis entity larger
                // TODO: Turn on 'Delete Successful Async Jobs'

                // Test 4
                //13830BE9-3E0C-E311-8B54-78E3B510FDBD
                EntityReference sCaseID =  (EntityReference)entity["new_case"];

                string sFetchXml1 = string.Format(@"
                <fetch distinct='false' mapping='logical' count='2' output-format='xml-platform' version='1.0'>
                  <entity name='new_serviceanalysis'>
                  	<filter>
                		<condition attribute='new_case' operator='eq' value='{0}'  />
                    </filter>
                	<order attribute='createdon' descending='true'/>
                  </entity>
                </fetch>", sCaseID.Id);

                FetchExpression query = new FetchExpression("new_serviceanalysis");
                query.Query = sFetchXml1;

                EntityCollection _saRecords = service.RetrieveMultiple(query);

                //Entity et = new Entity("new_serviceanalysis");

                switch (_saRecords.Entities.Count)
                {
                    case 0:
                        tracer.Trace("No Sales Analysis records found");
                        //return;
                        break;
                    case 1:
                        // Update To Stage & To Close for Latest Record
                        UpdateLatestRecord(serviceProvider, service, tracer, _saRecords);
                        break;
                    default:
                        // >=2
                        // Update To Stage & To Close for Latest Record
                        UpdateLatestRecord(serviceProvider, service, tracer, _saRecords);

                        // Update Older Record(s)
                        UpdateOlderRecord(serviceProvider, service, tracer, _saRecords);
                        break;
                }

                tracer.Trace("End of Plugin");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                "An error occurred in the plug-in.", ex);
            }
        }
        private void UpdateLatestRecord(IServiceProvider serviceProvider, IOrganizationService service, ITracingService tracer, EntityCollection _saRecords)
        {
            Entity record = _saRecords[0];
            Entity updateSA = new Entity("new_serviceanalysis");
            updateSA.Id = record.Id;

            DateTime createdOn = DateTime.MinValue;    //(DateTime)et["new_casecreated"];
            DateTime ownershipStart = DateTime.MinValue; //(DateTime)et["new_ownershipStart"];
            DateTime ownershipExitDate = DateTime.MinValue;    //(DateTime)et["new_ownershipexitdate"];
            decimal ownershipDays = 0; // DateTime.MinValue;  //(DateTime)et["new_ownershipdays"];

            if (record.Contains("new_casecreated")) { createdOn = (DateTime)record["new_casecreated"]; }
            if (record.Contains("new_ownershipstartdate")) { ownershipStart = (DateTime)record["new_ownershipstartdate"]; }
            if (record.Contains("new_ownershipexitdate")) { ownershipExitDate = (DateTime)record["new_ownershipexitdate"]; }
            if (record.Contains("new_ownershipdays")) { ownershipDays = (decimal)record["new_ownershipdays"]; }

            // If Days to Stage == NULL, Update it
            if (ownershipStart != DateTime.MinValue && createdOn != DateTime.MinValue)
            {
                updateSA["new_ownershipdays"] = GetTotalDays(createdOn, ownershipStart);
            }
            tracer.Trace(string.Format("Days to Stage: {0}", GetTotalDays(createdOn, ownershipStart)));

            // Always update the Days to Closed field
            if (ownershipExitDate != DateTime.MinValue && createdOn != DateTime.MinValue)
            {
                updateSA["new_daystoexit"] = GetTotalDays(createdOn, ownershipExitDate);
            }
            tracer.Trace(string.Format("Days to Close: {0}", GetTotalDays(createdOn, ownershipExitDate)));

            updateSA["new_tempstatus"] = "Plugin: UpdateLatestRecord";

            tracer.Trace("Updating Original Sales Analysis record");
            service.Update(updateSA);
            tracer.Trace("Update Successful :)");
        }

        private void UpdateOlderRecord(IServiceProvider serviceProvider, IOrganizationService service, ITracingService tracer, EntityCollection _saRecords)
        {
            Entity record = _saRecords[1];
            Entity updateSA = new Entity("new_serviceanalysis");
            updateSA.Id = record.Id;

            DateTime ownershipStart = DateTime.MinValue;
            if (record.Contains("new_ownershipStartDate")) { ownershipStart = (DateTime)record["new_ownershipStartDate"]; }

            DateTime exitOwnership = DateTime.Now.ToLocalTime();
            updateSA["new_ownershipexitdate"] = exitOwnership;
            updateSA["new_ownershipdays"] = GetTotalDays(ownershipStart, exitOwnership);

            updateSA["new_tempstatus"] = "Plugin: UpdateOlderRecord";

            tracer.Trace("Updating Original Sales Analysis record");
            service.Update(updateSA);
            tracer.Trace("Update Successful :)");
        }

        private Decimal GetTotalDays(DateTime dateFrom, DateTime dateTo)
        {
            TimeSpan ts = dateTo - dateFrom;
            Decimal d = new decimal(ts.TotalDays);

            return d;
        }
    }
}