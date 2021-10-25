// <copyright file="UpdateSPVProjectTask.cs" company="<company Name>">
//     2021 <Company Name>, All Rights Reserved.
// </copyright>
// File Name: UpdateSPVProjectTask.cs
// Description: This class will update the Project SPV on the Project Task if there is change in the Project Supervisor
//
/*
post update of project entity
This plugin will trigger on the post update of below mentioned fields
"ect_graphicssupervisor_lookup",
"ect_engineeringsupervisor_lookup",
"ect_programmingsupervisor_lookup",
"ect_projectlogisticsoperator_lookup",
"ect_digitaldeploymentspv_lookup",  
 */
// Created: June 09, 2021
// Author: Subhash Mahato, <Company Name>)
// Revisions:
// ===================================================================================
// VERSION         DATE            Modified By         DESCRIPTION
// -----------------------------------------------------------------------------------
// 1.0          06/09/2021          Subhash Mahato      CREATED
// 1.1          06/16/2021          Subhash Mahato      Updated: updated the namespace
// 1.2          08/03/2021          Subhash Mahato      Updated: Task 5914: Assignment for Teams on the project
// ===================================================================================

namespace SBT.ECT.Plugins.ProcessTemplateCopy
{
   using Microsoft.Xrm.Sdk;
   using Microsoft.Xrm.Sdk.Query;
   using SBT.ECT.Plugins.ProcessTemplateCopy.Helper;
   using System;
   using System.Collections.Generic;
   using System.Linq;

   public class UpdateSPVProjectTask : IPlugin
   {
      public void Execute(IServiceProvider serviceProvider)
      {
         IPluginExecutionContext context = null;
         IOrganizationService service = null;
         ITracingService tracingService = null;
         Entity project = null;

         context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
         IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
         service = serviceFactory.CreateOrganizationService(context.UserId);
         tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

         tracingService.Trace("context.MessageName. " + context.MessageName);

         // if this plugin is registered on any other entity then return
         if (context.PrimaryEntityName.ToLower() != "msdyn_project")
         {
            tracingService.Trace("plugin is expected to registered on the msdyn_project but registered on  " + context.PrimaryEntityName.ToLower());
            return;
         }

         string[] projectSupervisor = GetFieldList();

         try
         {
            switch (context.MessageName.ToLower())
            {
               case "update":
                  if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                  {
                     project = target;
                  }
                  break;
            }

            // return if the Type of the project is not template
            if (project.Contains("ect_istemplate") && project.GetAttributeValue<Boolean>("ect_istemplate") == true)
            {
               return;
            }

            // check if the above list is present in the given list
            var supervisors = projectSupervisor.Where(a => project.Contains(a));

            ProcessSupervisor(supervisors, project, service);

         }
         catch (InvalidPluginExecutionException)
         {
            throw;
         }
         catch (Exception)
         {
            throw;
         }
      }

      private static string[] GetFieldList()
      {
         return new string[] {
            "msdyn_projectmanager",
            "ect_graphicssupervisor_lookup",
            "ect_engineeringsupervisor_lookup",
            "ect_programmingsupervisor_lookup",
            "ect_projectlogisticsoperator_lookup",
            "ect_digitaldeploymentspv_lookup",
            "ect_plospvteam_lookup",
            "ect_psespvteam_lookup"
         };
      }

      private void ProcessSupervisor(IEnumerable<string> supervisors, Entity project, IOrganizationService service)
      {

         foreach (string supervisor in supervisors)
         {
            try
            {
               EntityCollection projectTasks = ReadTaskTOUpdate(supervisor, project, service);

               foreach (Entity projectTask in projectTasks.Entities)
               {
                  Entity OprojectTask = CreateProjecTaskObject(projectTask, project, supervisor);

                  service.Update(OprojectTask);
               }
            }
            catch (Exception ex)
            {
               GenerateSystemLog.CreateSystemLog(ex.Message, "ProcessSupervisor", project.LogicalName, "SBT.ECT.Plugins.ProcessTemplateCopy:UpdateSPVProjectTask", service);
            }
         }
      }

      private Entity CreateProjecTaskObject(Entity projectTask, Entity project, string supervisor)
      {
         Entity oProjectTask = new Entity(projectTask.LogicalName)
         {
            Id = projectTask.Id
         };
         oProjectTask["msdyn_project"] = new EntityReference("msdyn_project", project.Id);

         bool isUserrecordType = project.GetAttributeValue<EntityReference>(supervisor).LogicalName == "systemuser";
         string recordType = project.GetAttributeValue<EntityReference>(supervisor).LogicalName;

         if (projectTask.Contains("ae.ect_updateusertospvlookup_twooptions") && ((bool)((AliasedValue)projectTask["ae.ect_updateusertospvlookup_twooptions"]).Value) == true && isUserrecordType)
         {
            oProjectTask["ect_tasksupervisor_lookup"] = new EntityReference("systemuser", project.GetAttributeValue<EntityReference>(supervisor).Id);
         }

         if (projectTask.Contains("ae.ect_assignrecordtouser_twooptions") && ((bool)((AliasedValue)projectTask["ae.ect_assignrecordtouser_twooptions"]).Value) == true)
         {
            oProjectTask["ownerid"] = new EntityReference(recordType, project.GetAttributeValue<EntityReference>(supervisor).Id);
         }
         return oProjectTask;
      }

      private EntityCollection ReadTaskTOUpdate(string supervisor, Entity project, IOrganizationService service)
      {
         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                             <entity name='msdyn_projecttask'>
                               <attribute name='msdyn_subject' />
                               <attribute name='createdon' />
                               <attribute name='msdyn_projecttaskid' />
                               <attribute name='msdyn_project' />
                               <order attribute='msdyn_subject' descending='false' />
                               <filter type='and'>
                                 <condition attribute='msdyn_project' operator='eq' value='" + project.Id + @"' />
                               </filter>
                               <link-entity name='ect_taskcategory' from='ect_taskcategoryid' to='ect_taskcategory_lookup' link-type='inner' alias='ae'>
                                 <attribute name='ect_updateusertospvlookup_twooptions' />
                                 <attribute name='ect_assignrecordtouser_twooptions' />
                                 <filter type='and'>
                                   <condition attribute='ect_userschemanametocapture_text' operator='eq' value='" + supervisor + @"' />
                                 </filter>
                               </link-entity>
                             </entity>
                           </fetch>";

         EntityCollection projectTask = service.RetrieveMultiple(new FetchExpression(fetchString));
         return projectTask;
      }
   }
}
