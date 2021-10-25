// <copyright file="ProjectTaskAdjustment.cs" company="<company Name>">
//     2021 <Company Name>, All Rights Reserved.
// </copyright>
// File Name: ProjectTaskAdjustment.cs
// Description: 5912: set the Project Task date
// Created: January 13, 2021
// Author: Subhash Mahato, <Company Name>)
// Revisions:
// ===================================================================================
// VERSION         DATE            Modified By         DESCRIPTION
// -----------------------------------------------------------------------------------
// 1.0          08/25/2021          Subhash Mahato      CREATED
// ===================================================================================

namespace SBT.ECT.Plugins.ProcessTemplateCopy
{
   using Microsoft.Xrm.Sdk;
   using Microsoft.Xrm.Sdk.Query;
   using SBT.ECT.Plugins.ProcessTemplateCopy.Helper;
   using System;
   using System.Collections.Generic;
   using System.Linq;

   public class ProjectTaskAdjustment : IPlugin
   {
      public void Execute(IServiceProvider serviceProvider)
      {
         IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
         IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

         IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
         ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

         tracingService.Trace("context.MessageName: " + context.MessageName);

         if (context.PrimaryEntityName.ToLower() != "msdyn_project")
         {
            tracingService.Trace("Plugin is expected to registered on msdyn_projecttask but registered on " + context.PrimaryEntityName.ToLower());
            return;
         }

         try
         {
            Entity project = null;
            switch (context.MessageName.ToLower())
            {
               case "update":
                  if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                  {
                     project = target;
                  }
                  if (context.PostEntityImages.Contains("postImage") && context.PostEntityImages["postImage"] is Entity preImage)
                  {
                     project = preImage;
                  }
                  break;
            }

            // return if the Type of the project is not template
            if (project.Contains("ect_istemplate") && project.GetAttributeValue<Boolean>("ect_istemplate") == true)
            {
               tracingService.Trace($"Do not process the plugin if current project is template");
               return;
            }

            if (project.Contains("statuscode") && project.GetAttributeValue<OptionSetValue>("statuscode").Value != 1)
            {
               tracingService.Trace($"it is trigggered but the value is {project.GetAttributeValue<OptionSetValue>("statuscode").Value} it is returning");
               return;
            }

            OperationSetHandler.ProcessOperationSetOperation(service, tracingService, project.Id, $"{project.Id}:{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ss}",
               (operationSetId, processProjectServiceAPIRequest) => ProcessProjectTaskCreation(operationSetId, project, service));

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

      public void ProcessProjectTaskCreation(string operationSetId, Entity projectCurrent, IOrganizationService service)
      {
         if (projectCurrent != null && projectCurrent.Contains("ect_projecttemplateid"))
         {
            Guid templateProjectId = projectCurrent.GetAttributeValue<EntityReference>("ect_projecttemplateid").Id;
            Entity projectTemplate = ReadTemplateProject(templateProjectId, service);

            //1.  read the template resource assignment for the current task in the template
            //2.  get the team name and get the respective team name from the current proejct 
            ProcessProjectTask(projectCurrent, projectTemplate, operationSetId, service);
         }
      }

      private EntityCollection ReadCurrentProjectTasks(Guid projectId, IOrganizationService service)
      {
         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                   <entity name='msdyn_projecttask'>
                                     <attribute name='msdyn_subject' />
                                     <attribute name='createdon' />
                                     <attribute name='msdyn_projecttaskid' />
                                     <attribute name='msdyn_scheduledstart' />
                                    <attribute name='msdyn_duration' /> 
                                    <attribute name='ect_assoldversion_boolean' /> 
                                    <attribute name='ect_showinpsequoteform_twooptions' /> 
                                    <attribute name='ect_showinpqrformtasks_twooptions' /> 
                                    <attribute name='ect_showingqrformtasks_twooptions' /> 
                                    <attribute name='ect_isapprovalrequired_twooptions' />
			                           <attribute name='ect_showinmorformtasks_twooptions' /> 
                                    <attribute name='ect_showinddsrequestformtasks_twooptions' /> 
                                    <attribute name='ect_milestoneflags_gloptionset' /> 
                                    <attribute name='msdyn_parenttask' /> 
                                    <attribute name='ect_taskcategory_lookup' /> 
                                    <attribute name='ect_taskstatusreason_lookup' /> 
                                    <attribute name='ect_taskstatus_lookup' /> 
                                    <attribute name='msdyn_scheduledstart' /> 
                                    <attribute name='ect_automatedversiontype_gloptionset' /> 
			                           <attribute name='msdyn_description' />
		                              <attribute name='ect_notes' />
                                    <order attribute='msdyn_outlinelevel' descending='false' />
                                     <filter type='and'>
                                       <condition attribute='msdyn_project' operator='eq' value='" + projectId + @"' />
                                     </filter>
                                   </entity>
                                 </fetch>";

         return service.RetrieveMultiple(new FetchExpression(fetchString));
      }

      private Entity ReadTemplateProject(Guid currentProjectId, IOrganizationService service)
       => service.Retrieve("msdyn_project", currentProjectId, new ColumnSet(
          "ect_projecttemplateid",
          "msdyn_subject",
          "msdyn_scheduledstart"
          ));

      private static List<TaskCollection> GetProjectTaskList(EntityCollection projectTasks, EntityCollection templateProjectTasks)
      {
         var list = from currentTask in projectTasks.Entities
                    join templateTask in templateProjectTasks.Entities on currentTask.GetAttributeValue<string>("msdyn_subject") equals templateTask.GetAttributeValue<string>("msdyn_subject")
                    select new TaskCollection { oldValue = templateTask.Id.ToString(), newValue = currentTask.Id.ToString() };
         return list.ToList();
      }

      private void ProcessProjectTask(Entity currentProject, Entity projectTemplate, string operationSetId, IOrganizationService service)
      {
         ProcessProjectServiceAPIRequest processProjectServiceAPIRequest = new ProcessProjectServiceAPIRequest();

         EntityCollection currentProjectTasks = ReadCurrentProjectTasks(currentProject.Id, service);

         DateTime currentProjectStartDate = ReadProjectDate(currentProject);

         EntityCollection templateProjectTasks = ReadCurrentProjectTasks(projectTemplate.Id, service);

         List<TaskCollection> tlist = GetProjectTaskList(currentProjectTasks, templateProjectTasks);

         DateTime templateProjectStartDate = DateTime.MinValue;

         if (projectTemplate.Contains("msdyn_scheduledstart"))
         {
            templateProjectStartDate = projectTemplate.GetAttributeValue<DateTime>("msdyn_scheduledstart");
         }

         foreach (Entity projectTaskCreate in currentProjectTasks.Entities)
         {
            // check if the if the current task has resource allocation in the template 
            Guid taskId = Guid.Empty;
            List<Entity> singleTemplateTask = null;

            List<TaskCollection> getTemplateTask = tlist.Where(c => c.newValue == projectTaskCreate.Id.ToString()).ToList();

            if (getTemplateTask.Count > 0)
            {
               taskId = new Guid(getTemplateTask[0].oldValue);
               singleTemplateTask = templateProjectTasks.Entities.Where(c => c.Id == taskId).ToList();
            }
            if (singleTemplateTask.Count > 0 && templateProjectStartDate != DateTime.MinValue && currentProjectStartDate != DateTime.MinValue)
            {
               Entity projectTaskUpdate = GetTask(singleTemplateTask[0], projectTaskCreate.Id, tlist, currentProjectTasks, templateProjectStartDate, currentProjectStartDate);
               processProjectServiceAPIRequest.CallPssUpdateAction(projectTaskUpdate, operationSetId, service);
            }
         }
      }

      public Entity GetTask(Entity templateProjectTask, Guid currentProjectTaskId, List<TaskCollection> tlist, EntityCollection currentProjectTasks, DateTime templateProjectStartDate, DateTime currentProjectStartDate)
      {
         Entity projectTaskObject = new Entity("msdyn_projecttask", currentProjectTaskId);

         List<Entity> getRecordWith = null;
         // this will check if this is a parent task, if yes than do not update any value as it calculate autometically
         getRecordWith = currentProjectTasks.Entities.Where(c => (c.GetAttributeValue<EntityReference>("msdyn_parenttask") != null && c.GetAttributeValue<EntityReference>("msdyn_parenttask").Id == currentProjectTaskId)).ToList();

         // do not set the system field on the parent record
         if (getRecordWith == null || getRecordWith.Count == 0)
         {
            projectTaskObject["msdyn_duration"] = templateProjectTask.GetAttributeValue<double>("msdyn_duration");

            if (templateProjectTask.GetAttributeValue<DateTime>("msdyn_scheduledstart") != DateTime.MinValue)
            {
               double daysDifference = (templateProjectTask.GetAttributeValue<DateTime>("msdyn_scheduledstart") - templateProjectStartDate).TotalDays;
               DateTime startDate = currentProjectStartDate.AddDays(daysDifference);
               projectTaskObject["msdyn_scheduledstart"] = startDate;
            }
         }
         return projectTaskObject;
      }

      private DateTime ReadProjectDate(Entity currentProject)
      {
         DateTime currentProjectStartDate = DateTime.MinValue;
         if (currentProject.Contains("msdyn_scheduledstart") && currentProject.GetAttributeValue<DateTime>("msdyn_scheduledstart") != null)
         {
            currentProjectStartDate = currentProject.GetAttributeValue<DateTime>("msdyn_scheduledstart");
         }
         return currentProjectStartDate;
      }
   }
}

public class TaskCollection
{
   public string oldValue;

   public string newValue;
}
