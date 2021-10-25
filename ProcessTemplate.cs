// <copyright file="ProcessTemplate.cs" company="<company Name>">
//     2021 <Company Name>, All Rights Reserved.
// </copyright>
// File Name: ProcessTemplate.cs
// Description: This class will create copy functionality of the template
// Created: January 13, 2021
// Author: Subhash Mahato, <Company Name>)
// Revisions:
// ===================================================================================
// VERSION         DATE            Modified By         DESCRIPTION
// -----------------------------------------------------------------------------------
// 1.0          13/01/2021          Subhash Mahato      CREATED
// 1.1          01/02/2021          Subhash Mahato      Added the logic to create task from template using project service API
//                                                      but it still have issues like parenting which is not provided in the current api, so not implemented 
// 1.2          03/08/2021          Subhash Mahato      Added the logic to update the RFI expected days on the Project
// 1.3          03/11/2021          Subhash Mahato      commented the logic which throws the error if there is no Template present
// 1.4          03/15/2021          Subhash Mahato      Update the optionset value to read the master configuration of the RFI type (RFI Expected Time )  402220002
// 1.5          04/15/2021          Subhash Mahato      Updated 1- Update the SPV lookup fields on the project form: check the zone supervisors, if they contain data when update the related
//                                                                 supervisor lookups on the project with the related people. 
//                                                                 If the supervisor lookups on the zone do not contain data, then check the sales office to see the supervisor lookups and update the 
//                                                                 project with the related supervisor lookup users. This checking can be made on SPV level. For example, if the Graphics supervisor 
//                                                                 lookup is blank on the Zone, then check the Graphics SPV on the sales office.   
// 1.6          04/15/2021          Subhash Mahato      Updated: Need add the supervisior from the two seprate logic GOC -General and NOC- Government also need to read first from zone and than salesoffice is it is not present
// 1.7          05/14/2021          Subhash Mahato      Update: Need to update the logic for the add ZTC SPV Engineering to the Settings/Zones (5370)
//                                                      Need to check the business type in the ect configuration for the business type ect (C0014), there need to verify the Ect SPV Assignment.
// 1.8          05/19/2021          Subhash Mahato      Update: added the logic to set the PLO and DD SPV
// 1.9          07/20/2021          Subhash Mahato      Update: Task 5803: change in the logic of programming SPV and ZTC Eng SPV assignment in project opening
// 2.0          08/03/2021          Subhash Mahato      Update: Task 5914: Assignment for Teams on the project
// ===================================================================================

namespace SBT.ECT.Plugins.ProcessTemplateCopy
{
   using Microsoft.Xrm.Sdk;
   using Microsoft.Xrm.Sdk.Query;
   using SBT.ECT.Plugins.ProcessTemplateCopy.Helper;
   using System;
   using System.Collections.Generic;
   using System.Linq;

   public class ProcessTemplate : IPlugin
   {
      public void Execute(IServiceProvider serviceProvider)
      {
         Entity project = null;
         IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
         IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
         IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
         ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

         tracingService.Trace("context.MessageName" + context.MessageName);

         // if this plugin is registered on any other entity then return         
         if (context.PrimaryEntityName.ToLower() != "msdyn_project")
         {
            tracingService.Trace($"plugin was expected to be registered on msdyn_project but registered on " + context.PrimaryEntityName.ToLower());
            return;
         }

         /*
          *  Make sure that the no operation is performed on project if type is template .
          */

         /*
          *  1. Plugin to check the configurations table, 
          *  2. Check the configurations with the type: template, query with the sales amount with the min-max values on the configuration, 
          *  3. if <50 pick the project on the configuraiton; 
          *  4. if =>50 check the business type and pick the template that fits. 
          */

         try
         {
            switch (context.MessageName.ToLower())
            {
               case "create":
                  if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
                  {
                     project = entity;
                  }
                  break;
            }

            // return if the Type of the project is not template
            if (project.Contains("ect_istemplate") && project.GetAttributeValue<Boolean>("ect_istemplate") == true)
            {
               tracingService.Trace($"Do not process the plugin if current project is template");
               return;
            }

            // Set Template on Current Project
            ProcessCurrentTemplate(project, service);

            ProcessTCSupervisor(project, project.Id, service);

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
      /// <summary>
      /// Plugin to check the configurations table, check the configurations with the type: template, query with the sales amount with the min-max values on the configuration, 
      /// "if <50 pick the project on the configuraiton; 
      /// if =>50 check the business type and pick the template that fits.
      /// </summary>
      /// <param name="entity">.</param>
      /// <param name="service">The service<see cref="IOrganizationService"/>.</param>
      private void ProcessCurrentTemplate(Entity entity, IOrganizationService service)
      {
         decimal saleAmount = 0;
         Guid businessType = Guid.Empty;
         Entity projectTemplate = null;

         // read the curernt value
         if (entity.Contains("ect_salesamount") && entity.GetAttributeValue<Money>("ect_salesamount") != null)
         {
            saleAmount = entity.GetAttributeValue<Money>("ect_salesamount").Value;
         }

         // business Type
         if (entity.Contains("ect_businesstypeid") && entity.GetAttributeValue<EntityReference>("ect_businesstypeid") != null)
         {
            businessType = entity.GetAttributeValue<EntityReference>("ect_businesstypeid").Id;
         }

         //Read the project configuration
         EntityCollection projectConfiguration = ReadProjectConfiguration(service);

         // if this not null
         if (projectConfiguration != null)
         {
            projectTemplate = SelectCorrectTempate(projectConfiguration, saleAmount, businessType);
         }

         Guid projectId = entity.Id;
         // Read the Project task for the given task
         if (projectTemplate != null && projectTemplate.Contains("ect_projectid"))
         {
            Guid projectTemplateId = projectTemplate.GetAttributeValue<EntityReference>("ect_projectid").Id;
            EntityCollection projectTasks = ReadProjectTemplateTask(projectTemplateId, service);

            // update the template on project
            UpdateProject(projectId, projectTemplateId, service);
         }

         /**
          * 1. Read the ect configuration with the configuration type is "RFI Expected Time"
          * 2. Update the RFI value from Number days to the RFI Expected Response Time in Projects
          */
         Entity rfiConfiguration = ReadRFiConfiguration(service);

         if (rfiConfiguration != null)
         {
            UpdateProjectRFi(projectId, rfiConfiguration, service);
         }
      }

      private void UpdateProjectRFi(Guid projectId, Entity rfiConfiguration, IOrganizationService service)
      {
         // tracingService.Trace("projectId: " + projectId);
         Entity curerntProject = new Entity
         {
            LogicalName = "msdyn_project",
            Id = projectId
         };
         curerntProject["ect_rfiexpectedresponsetimeday_wholenumber"] = rfiConfiguration.Contains("ect_numberofdays_wholenumber") == true ?
            rfiConfiguration.GetAttributeValue<Int32>("ect_numberofdays_wholenumber") : 0;
         service.Update(curerntProject);
      }

      private void UpdateProject(Guid projectId, Guid projectTemplateId, IOrganizationService service)
      {
         // tracingService.Trace("projectId 1: " + projectId);
         Entity curerntProject = new Entity
         {
            LogicalName = "msdyn_project",
            Id = projectId
         };
         curerntProject["ect_projecttemplateid"] = new EntityReference("msdyn_project", projectTemplateId);
         //curerntProject["statuscode"] = new OptionSetValue(192350001);
         service.Update(curerntProject);
      }

      private EntityCollection ReadProjectTemplateTask(Guid projectId, IOrganizationService service)
      {
         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_projecttask'>
                                <all-attributes />
                                <order attribute='msdyn_transactioncategory' descending='false' />
                                <filter type='and'>
                                  <condition attribute='msdyn_project' operator='eq' value='" + projectId + @"' />
                                </filter>
                              </entity>
                            </fetch>";

         EntityCollection projectTasks = service.RetrieveMultiple(new FetchExpression(fetchString));
         return projectTasks;
      }
      /// <summary>
      /// This functionn will give the correct template based on the value provided.
      /// </summary>
      /// <param name="projectConfiguration"></param>
      /// <param name="saleAmount"></param>
      /// <param name="businessType"></param>
      /// <returns></returns>
      private Entity SelectCorrectTempate(EntityCollection projectConfiguration, decimal saleAmount, Guid businessType)
      {
         Entity defaultTemplate = null;
         List<Entity> projectConfigurationDefault = new List<Entity>();

         // parse it to the List
         List<Entity> entityList = projectConfiguration.Entities.ToList();

         // 1. Read the default record
         projectConfigurationDefault = entityList.Where(x => x.GetAttributeValue<Money>("ect_minimumvalue").Value <= saleAmount
         &&
         saleAmount <= x.GetAttributeValue<Money>("ect_maximumvalue").Value).ToList();

         // 2. if record count is one than set this record
         if (projectConfigurationDefault.Count == 1)
         {
            Entity templateRecord = projectConfigurationDefault[0];

            // if business type contains value than try to compare the business type with the given result
            if (templateRecord.Contains("ect_businesstypeid"))
            {
               projectConfigurationDefault = entityList.Where(x => (
               x.GetAttributeValue<Money>("ect_minimumvalue").Value <= saleAmount
               &&
               saleAmount <= x.GetAttributeValue<Money>("ect_maximumvalue").Value
               &&
               x.GetAttributeValue<EntityReference>("ect_businesstypeid").Id.Equals(businessType))).ToList();

               if (projectConfigurationDefault.Count == 1)
               {
                  defaultTemplate = projectConfigurationDefault[0];
                  //return defaultTemplate;
               }
            }
            else
            {
               defaultTemplate = templateRecord;
            }
         }
         else if (projectConfigurationDefault.Count > 1)
         {
            // 3. if the record count is more than one than try to pass the business type as well,               
            projectConfigurationDefault = entityList.Where(x => (x.GetAttributeValue<Money>("ect_minimumvalue").Value <= saleAmount
            &&
            saleAmount <= x.GetAttributeValue<Money>("ect_maximumvalue").Value
            &&
            x.GetAttributeValue<EntityReference>("ect_businesstypeid").Id.Equals(businessType))).ToList();

            if (projectConfigurationDefault.Count == 1)
            {
               defaultTemplate = projectConfigurationDefault[0];
            }
         }
         return defaultTemplate;
      }
      /// <summary>
      /// 1- check the sales office details. 
      /// If there is fire-security-automation TC, pick the appropriate one for the business type of the project. 
      /// If only the Default TC is filled, then pick default TC. 
      /// If no definitions on the Sales office, then pick the TC on the Zone. 
      /// If none, then give warning - no TC definitions exist or the zone or the sales office.
      /// </summary>
      /// <param name="projectInfo">The projectInfo<see cref="Entity"/>.</param>
      /// <param name="projectId">The projectId<see cref="Guid"/>.</param>
      /// <param name="service">The service<see cref="IOrganizationService"/>.</param>
      private void ProcessTCSupervisor(Entity projectInfo, Guid projectId, IOrganizationService service)
      {
         Entity salesOffice = null;
         Entity zone = null;
         Entity spvConfiguration = null;
         bool isTechnicalCoordinatorPresent = false;

         /*
          * Once the user fills the Zone and Sales office attributes and saves the project afterwards, the system will do the below: 
             1- check the TC definitions on the sales office. If any of fire-security-automation TC lookups include data, then check the business type of the project. If the business type = fire and fire TC contains data, select the fire TC on Sales Office. the same applies for security and automation as well. 
             2- If the fire-automation-security TC lookups do not contain data or the filled lookup does not match with the business type of the project, then check if the Default TC lookup on the Sales office contains data. If yes, then take that calue to the project TC lookup. 
             3- If the default TC on the sales office does not contain data, then take the default TC lookup value on the Zone to the Project Technical Coordinator field. 
             4- If none of these contains data, then throw an error stating: "There are no TC definitions made for the sales office or the zone. Please make the TC entry manually." and make the TC field editable until someone chooses a user on the form and saves it. the field should still be editable until the form is saved so that if the user makes any mistake, can correct the mistake. 
          */

         Guid saleOfficeId = Guid.Empty;
         // read the curernt value
         if (projectInfo.Contains("ect_salesofficeid") && projectInfo.GetAttributeValue<EntityReference>("ect_salesofficeid") != null)
         {
            saleOfficeId = projectInfo.GetAttributeValue<EntityReference>("ect_salesofficeid").Id;
         }

         Guid zoneId = Guid.Empty;
         if (projectInfo.Contains("ect_zoneid") && projectInfo.GetAttributeValue<EntityReference>("ect_zoneid") != null)
         {
            zoneId = projectInfo.GetAttributeValue<EntityReference>("ect_zoneid").Id;
         }

         bool isNOC = false;
         if (projectInfo.Contains("ect_federaljobflag_twooptions"))
         {
            isNOC = projectInfo.GetAttributeValue<bool>("ect_federaljobflag_twooptions");
         }

         Entity businessType = null;
         string businessTypeName = string.Empty;
         // read the business type name
         if (projectInfo.Contains("ect_businesstypeid") && projectInfo.GetAttributeValue<EntityReference>("ect_businesstypeid") != null)
         {
            Guid businessTypeNameId = projectInfo.GetAttributeValue<EntityReference>("ect_businesstypeid").Id;
            businessType = ReadBusinessType(service, businessTypeNameId);
            businessTypeName = businessType.GetAttributeValue<string>("ect_businesstype").ToLower();
         }

         Guid technicalCoordination = Guid.Empty;
         // read the sales office
         if (saleOfficeId != Guid.Empty)
         {
            salesOffice = ReadSalesOffices(service, saleOfficeId);
            technicalCoordination = ProcessSalesOfficeRecord(salesOffice, businessType);
         }

         if (zoneId != Guid.Empty)
         { zone = ReadZone(service, zoneId); }

         if (projectInfo != null && projectInfo.Contains("ect_businesstypeid") && projectInfo.GetAttributeValue<EntityReference>("ect_businesstypeid") != null)
         {
            // read the proejct configuration 
            spvConfiguration = ReadSPVConfiguration(projectInfo.GetAttributeValue<EntityReference>("ect_businesstypeid").Id, service);
         }

         bool isZTCPresent = false;
         if (spvConfiguration != null)
         {
            isZTCPresent = true;
         }

         // check if the tenant is present or not
         if (technicalCoordination == Guid.Empty && zoneId != Guid.Empty)
         {
            technicalCoordination = ReadTechincalCoordinator(service, zoneId);
         }

         if (technicalCoordination != Guid.Empty)
         {
            isTechnicalCoordinatorPresent = true;
         }

         Guid projectLogiceOfficer = Guid.Empty;
         Guid digitalDeploymentSPV = Guid.Empty;
         Guid programmingSupervisor = Guid.Empty;
         Guid graphicsSupervisor = Guid.Empty;
         Guid engineeringSupervisor = Guid.Empty;
         EntityReference zonePLOSPVTeam = null;
         EntityReference zonePSESPVTeam = null;
         if (zone != null && salesOffice != null)
         {
            if (salesOffice.Contains("ect_plo_lookup"))
            {
               projectLogiceOfficer = salesOffice.GetAttributeValue<EntityReference>("ect_plo_lookup").Id;
            }
            if (zone.Contains("ect_digitaldeploymentspv_lookup"))
            {
               digitalDeploymentSPV = zone.GetAttributeValue<EntityReference>("ect_digitaldeploymentspv_lookup").Id;
            }

            ProcessSupervisor(salesOffice,
               zone,
               isNOC,
               ref programmingSupervisor,
               ref graphicsSupervisor,
               ref engineeringSupervisor,
               isZTCPresent,
               businessTypeName
               );


            ReadZoneTeam(zone, isNOC, ref zonePLOSPVTeam, ref zonePSESPVTeam);
         }

         UpdateProjectTCSupervisor(technicalCoordination,
            projectId,
            isTechnicalCoordinatorPresent,
            programmingSupervisor,
            graphicsSupervisor,
            engineeringSupervisor,
            projectLogiceOfficer,
            digitalDeploymentSPV,
            zonePLOSPVTeam,
            zonePSESPVTeam,
            service
            );
      }

      private void ReadZoneTeam(Entity zone, bool isNOC, ref EntityReference zonePLOSPVTeam, ref EntityReference zonePSESPVTeam)
      {
         //"ect_nocpsespvteam_lookup",
         //"ect_gocpsespvteam_lookup",

         if (zone.Contains("ect_nocpsespvteam_lookup") && isNOC == true)
         {
            zonePSESPVTeam = zone.GetAttributeValue<EntityReference>("ect_nocpsespvteam_lookup");
         }
         else if (zone.Contains("ect_gocpsespvteam_lookup") && isNOC == false)
         {
            zonePSESPVTeam = zone.GetAttributeValue<EntityReference>("ect_gocpsespvteam_lookup");
         }

         if (zone.Contains("ect_zoneplospvteam_lookup"))
         {
            zonePLOSPVTeam = zone.GetAttributeValue<EntityReference>("ect_zoneplospvteam_lookup");
         }
      }

      private static Entity ReadZone(IOrganizationService service, Guid zoneId)
      {
         return service.Retrieve("ect_zone", zoneId, new ColumnSet("ect_defaulttcid",
            "ect_nocspvprogramming_lookup",
            "ect_nocspvgraphics_lookup",
            "ect_nocspvengineering_lookup",
            "ect_gocspvprogramming_lookup",
            "ect_gocspvgraphics_lookup",
            "ect_digitaldeploymentspv_lookup",
            "ect_gocprogrammingspvautomation",
            "ect_gocprogrammingspvfire_lookup",
            "ect_gocprogrammingspvsecurity_lookup",
            "ect_nocpsespvteam_lookup",
            "ect_gocpsespvteam_lookup",
            "ect_zoneplospvteam_lookup",
            "ect_nocprogrammingspvautomation_lookup",
            "ect_nocprogrammingspvfire_lookup",
            "ect_nocprogrammingspvsecurity_lookup",
            "ect_ztcengineeringspvautomation_lookup",
            "ect_ztcengineeringspvfire_lookup",
            "ect_ztcengineeringspvsecurity_lookup"
            ));
      }
      /// <summary>
      /// Read the Project configuration of SPV Assignment this record wich has confgiured as.
      /// </summary>
      /// <param name="businessType">.</param>
      /// <param name="service">The service<see cref="IOrganizationService"/>.</param>
      /// <returns>The <see cref="Entity"/>.</returns>
      private Entity ReadSPVConfiguration(Guid businessType, IOrganizationService service)
      {
         Entity spvConfiguration = null;
         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='ect_configuration'>
                            <attribute name='ect_configurationname' />
                            <attribute name='ect_configurationtype' />
                            <attribute name='ect_businesstypeid' />
                            <attribute name='ect_configurationid' />
                            <attribute name='ect_engineeringspvassignment_optionset' />
                            <order attribute='ect_configurationname' descending='false' />
                            <filter type='and'>
                              <condition attribute='ect_businesstypeid' operator='eq' uiname='Fire' uitype='ect_businesstype' value='" + businessType + @"' />
                              <condition attribute='ect_configurationcode_text' operator='eq' value='c0014' />
                              <condition attribute='ect_engineeringspvassignment_optionset' operator='eq' value='402220000' />
                            </filter>
                          </entity>
                        </fetch>";
         EntityCollection rfiConfigurations = service.RetrieveMultiple(new FetchExpression(fetchString));
         if (rfiConfigurations.Entities.Count >= 1)
         {
            spvConfiguration = rfiConfigurations.Entities[0];
         }
         return spvConfiguration;
      }
      /// <summary>
      /// this function will return the various supervisior based on the zone/region.
      /// </summary>
      /// <param name="salesOffice"></param>
      /// <param name="zone"></param>
      /// <param name="isNOC"></param>
      /// <param name="programmingSupervisor"></param>
      /// <param name="graphicsSupervisor"></param>
      /// <param name="engineeringSupervisor"></param>
      /// <param name="isZTCPresent"></param>
      /// <param name="businessTypeName"></param>
      private void ProcessSupervisor(Entity salesOffice,
     Entity zone,
     bool isNOC,
     ref Guid programmingSupervisor,
     ref Guid graphicsSupervisor,
     ref Guid engineeringSupervisor,
     bool isZTCPresent,
     string businessTypeName)
      {
         // first check if the supervisor present in the sales office only than go for the zone
         /* logic update Need add the supervisior from the two seprate logic GOC -General and NOC- Government */


         if (isNOC) /* this for NOC*/
         {
            /*1*/
            programmingSupervisor = GetProgrammingSupervisor(salesOffice, zone, businessTypeName, isNOC);

            /*2*/

            if (salesOffice.Contains("ect_nocspvgraphics_lookup"))
            {
               graphicsSupervisor = salesOffice.GetAttributeValue<EntityReference>("ect_nocspvgraphics_lookup").Id;
            }
            else if (zone.Contains("ect_nocspvgraphics_lookup"))
            {
               graphicsSupervisor = zone.GetAttributeValue<EntityReference>("ect_nocspvgraphics_lookup").Id;
            }

            //Zone PLO SPV Team

         }
         else if (isNOC == false) /*this is for the GOC general */
         {
            /*1*/
            programmingSupervisor = GetProgrammingSupervisor(salesOffice, zone, businessTypeName, isNOC);

            /*2*/

            if (salesOffice.Contains("ect_gocspvgraphics_lookup"))
            {
               graphicsSupervisor = salesOffice.GetAttributeValue<EntityReference>("ect_gocspvgraphics_lookup").Id;
            }
            else if (zone.Contains("ect_gocspvgraphics_lookup"))
            {
               graphicsSupervisor = zone.GetAttributeValue<EntityReference>("ect_gocspvgraphics_lookup").Id;
            }

         }

         ///*3*/
         if (isZTCPresent == false)
         {
            engineeringSupervisor = GetEngineeringSupervisor(salesOffice, zone, businessTypeName, isNOC);
         }

         // this replace the engineering with value present in the ztc engineering looks as this is the priority
         if (isZTCPresent)
         {
            engineeringSupervisor = ReadZTEEngineering(salesOffice, zone, engineeringSupervisor, businessTypeName);
         }
      }
      /// <summary>
      /// this function will read the data from zte
      /// </summary>
      /// <param name="salesOffice"></param>
      /// <param name="zone"></param>
      /// <param name="engineeringSupervisor"></param>
      /// <param name="businessTypeName"></param>
      /// <returns></returns>
      private static Guid ReadZTEEngineering(Entity salesOffice, Entity zone, Guid engineeringSupervisor, string businessTypeName)
      {
         switch (businessTypeName)
         {
            case "automation":
               engineeringSupervisor = GetEngineeringSupervisor(salesOffice, zone, "ect_ztcengineeringspvautomation_lookup");
               break;

            case "fire":
               engineeringSupervisor = GetEngineeringSupervisor(salesOffice, zone, "ect_ztcengineeringspvfire_lookup");
               break;

            case "security":
               engineeringSupervisor = GetEngineeringSupervisor(salesOffice, zone, "ect_ztcengineeringspvsecurity_lookup");
               break;
         }

         return engineeringSupervisor;
      }
      /// <summary>
      /// this function will get the engineering supervisor.
      /// </summary>
      /// <param name="salesOffice"></param>
      /// <param name="zone"></param>
      /// <param name="fieldName"></param>
      /// <returns></returns>
      private static Guid GetEngineeringSupervisor(Entity salesOffice, Entity zone, string fieldName)
      {
         Guid engineeringSupervisor = Guid.Empty;
         if (salesOffice.Contains(fieldName))
         {
            engineeringSupervisor = salesOffice.GetAttributeValue<EntityReference>(fieldName).Id;
         }
         else if (zone.Contains(fieldName))
         {
            engineeringSupervisor = zone.GetAttributeValue<EntityReference>(fieldName).Id;
         }
         return engineeringSupervisor;
      }
      /// <summary>
      /// this function will retunr the engineering based on the business type.
      /// </summary>
      /// <param name="salesOffice">.</param>
      /// <param name="zone">.</param>
      /// <param name="businessTypeName">.</param>
      /// <param name="isNOC">The isNOC<see cref="bool"/>.</param>
      private static Guid GetEngineeringSupervisor(Entity salesOffice, Entity zone, string businessTypeName, bool isNOC)
      {
         Guid engineeringSupervisor = Guid.Empty;
         switch (isNOC)
         {
            case true:

               if (Constant.BUSINESSTYPEAUTOMATION == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_nocengineeringspvautomation_lookup");
               }
               else if (Constant.BUSINESSTYPEFIRE == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_nocengineeringspvfire_lookup");
               }
               else if (Constant.BUSINESSTYPESECURITY == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_nocengineeringspvsecurity_lookup");
               }

               break;
            case false:

               if (Constant.BUSINESSTYPEAUTOMATION == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_gocengineeringspvautomation_lookup");
               }
               else if (Constant.BUSINESSTYPEFIRE == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_gocengineeringspvfire_lookup");
               }
               else if (Constant.BUSINESSTYPESECURITY == businessTypeName)
               {
                  engineeringSupervisor = GetSupervisor(salesOffice, zone, "ect_gocengineeringspvsecurity_lookup");
               }
               break;
         }

         return engineeringSupervisor;
      }

      private static Guid GetSupervisor(Entity salesOffice, Entity zone, string fieldName)
      {
         Guid engineeringSupervisor = Guid.Empty;

         if (salesOffice.Contains(fieldName))
         {
            engineeringSupervisor = salesOffice.GetAttributeValue<EntityReference>(fieldName).Id;
         }

         else if (zone.Contains(fieldName))
         {
            engineeringSupervisor = zone.GetAttributeValue<EntityReference>(fieldName).Id;
         }

         return engineeringSupervisor;
      }
      /// <summary>
      /// this function will retunr the engineering based on the business type.
      /// </summary>
      /// <param name="salesOffice">.</param>
      /// <param name="zone">.</param>
      /// <param name="businessTypeName">.</param>
      /// <param name="isNOC">The isNOC<see cref="bool"/>.</param>
      /// <returns>.</returns>
      private static Guid GetProgrammingSupervisor(Entity salesOffice, Entity zone, string businessTypeName, bool isNOC)
      {
         Guid programmingSupervisor = Guid.Empty;
         switch (isNOC)
         {
            case true:

               if (Constant.BUSINESSTYPEAUTOMATION == businessTypeName)
               {
                  programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_nocprogrammingspvautomation_lookup");
               }
               else if (Constant.BUSINESSTYPEFIRE == businessTypeName)
               {
                  programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_nocprogrammingspvfire_lookup");
               }
               else if (Constant.BUSINESSTYPESECURITY == businessTypeName)
               {
                  programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_nocprogrammingspvsecurity_lookup");
               }

               break;
            case false:

               if (Constant.BUSINESSTYPEAUTOMATION == businessTypeName)
               {
                  if (salesOffice.Contains("ect_gocprogrammingspvautomation_lookup"))
                  {
                     programmingSupervisor = salesOffice.GetAttributeValue<EntityReference>("ect_gocprogrammingspvautomation_lookup").Id;
                  }
                  else if (zone.Contains("ect_gocprogrammingspvautomation"))
                  {
                     programmingSupervisor = zone.GetAttributeValue<EntityReference>("ect_gocprogrammingspvautomation").Id;
                  }
                  //programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_gocprogrammingspvautomation_lookup");

               }
               else if (Constant.BUSINESSTYPEFIRE == businessTypeName)
               {
                  programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_gocprogrammingspvfire_lookup");
               }
               else if (Constant.BUSINESSTYPESECURITY == businessTypeName)
               {
                  programmingSupervisor = ReadProgrammingSupervisor(salesOffice, zone, "ect_gocprogrammingspvsecurity_lookup");
               }
               break;
         }

         return programmingSupervisor;
      }

      private static Guid ReadProgrammingSupervisor(Entity salesOffice, Entity zone, string fieldName)
      {
         Guid programmingSupervisor = Guid.Empty;
         if (salesOffice.Contains(fieldName))
         {
            programmingSupervisor = salesOffice.GetAttributeValue<EntityReference>(fieldName).Id;
         }
         else if (zone.Contains(fieldName))
         {
            programmingSupervisor = zone.GetAttributeValue<EntityReference>(fieldName).Id;
         }

         return programmingSupervisor;
      }

      private Guid ReadTechincalCoordinator(IOrganizationService service, Guid zoneId)
      {
         Guid TenantCoordination = Guid.Empty;

         Entity zone = service.Retrieve("ect_zone", zoneId, new ColumnSet("ect_defaulttcid",
   "ect_nocspvprogramming_lookup",
   "ect_nocspvgraphics_lookup",
   "ect_nocspvengineering_lookup",
   "ect_gocspvprogramming_lookup",
   "ect_gocspvgraphics_lookup"

   ));
         if (zone != null && zone.Contains("ect_defaulttcid"))
         {
            TenantCoordination = zone.GetAttributeValue<EntityReference>("ect_defaulttcid").Id;
         }

         return TenantCoordination;
      }

      private void UpdateProjectTCSupervisor(Guid technicaltCoordination,
     Guid projectId,
     bool isTechnicalCoordinatorPresent,
     Guid programmingSupervisor,
     Guid graphicsSupervisor,
     Guid engineeringSupervisor,
     Guid projectLogiceOfficer,
     Guid digitalDeploymentSPV,
     EntityReference zonePLOSPVTeam,
     EntityReference zonePSESPVTeam,
     IOrganizationService service)
      {
         Entity curerntProject = new Entity
         {
            LogicalName = "msdyn_project",
            Id = projectId
         };

         if (technicaltCoordination != Guid.Empty)
         {
            curerntProject["ect_technicalcoordinatorid"] = new EntityReference("systemuser", technicaltCoordination);
         }

         if (programmingSupervisor != Guid.Empty)
         {
            curerntProject["ect_programmingsupervisor_lookup"] = new EntityReference("systemuser", programmingSupervisor);
         }

         if (graphicsSupervisor != Guid.Empty)
         {
            curerntProject["ect_graphicssupervisor_lookup"] = new EntityReference("systemuser", graphicsSupervisor);
         }

         if (engineeringSupervisor != Guid.Empty)
         {
            curerntProject["ect_engineeringsupervisor_lookup"] = new EntityReference("systemuser", engineeringSupervisor);
         }

         if (projectLogiceOfficer != Guid.Empty)
         {
            curerntProject["ect_projectlogisticsoperator_lookup"] = new EntityReference("systemuser", projectLogiceOfficer);
         }

         if (digitalDeploymentSPV != Guid.Empty)
         {
            curerntProject["ect_digitaldeploymentspv_lookup"] = new EntityReference("systemuser", digitalDeploymentSPV);
         }

         if (zonePLOSPVTeam != null)
         {
            curerntProject["ect_plospvteam_lookup"] = zonePLOSPVTeam;
         }

         if (zonePSESPVTeam != null)
         {
            curerntProject["ect_psespvteam_lookup"] = zonePSESPVTeam;
         }

         curerntProject["ect_istcpresent"] = isTechnicalCoordinatorPresent;

         service.Update(curerntProject);
      }
      /// <summary>
      /// Read the TC id from the service office.
      /// </summary>
      /// <param name="salesOffice">The salesOffice<see cref="Entity"/>.</param>
      /// <param name="businessType">.</param>
      /// <returns>.</returns>
      private static Guid ProcessSalesOfficeRecord(Entity salesOffice, Entity businessType)
      {

         Guid TechnicalCoordination = Guid.Empty;

         bool isAnyTcContainsValue = false;
         string businessTypeName = string.Empty;

         if ((businessTypeName == null || businessTypeName == "") && businessType != null)
         {
            businessTypeName = businessType.GetAttributeValue<string>("ect_businesstype");
         }

         // check if the sales office is not null
         if (salesOffice != null && (salesOffice.Contains("ect_firetcid") || salesOffice.Contains("ect_automationtcid") || salesOffice.Contains("ect_securitytcid")))
         {
            isAnyTcContainsValue = true;
         }

         if (isAnyTcContainsValue && businessTypeName.ToLower() == "fire" && salesOffice.Contains("ect_firetcid"))
         {
            TechnicalCoordination = salesOffice.GetAttributeValue<EntityReference>("ect_firetcid").Id;
         }
         else if (isAnyTcContainsValue && businessTypeName.ToLower() == "automation" && salesOffice.Contains("ect_automationtcid"))
         {
            TechnicalCoordination = salesOffice.GetAttributeValue<EntityReference>("ect_automationtcid").Id;
         }
         else if (isAnyTcContainsValue && businessTypeName.ToLower() == "security" && salesOffice.Contains("ect_securitytcid"))
         {
            TechnicalCoordination = salesOffice.GetAttributeValue<EntityReference>("ect_securitytcid").Id;
         }
         else if (salesOffice.Contains("ect_defaulttcid"))
         {
            TechnicalCoordination = salesOffice.GetAttributeValue<EntityReference>("ect_defaulttcid").Id;
         }

         return TechnicalCoordination;
      }

      private static Entity ReadSalesOffices(IOrganizationService service, Guid saleOfficeId)
      {
         // read the sales office 
         return service.Retrieve("ect_salesoffice", saleOfficeId, new ColumnSet("ect_salesofficeid",
             "ect_salesoffice",
             "ect_zoneid",
             "ect_securitytcid",
             "ect_salesofficecode",
             "ect_firetcid",
             "ect_defaulttcid",
             "ect_automationtcid",
             "ect_nocspvprogramming_lookup",
             "ect_nocspvgraphics_lookup",
             "ect_nocspvengineering_lookup",
             "ect_gocspvprogramming_lookup",
             "ect_gocspvgraphics_lookup",
             "ect_ztcspvengineering_lookup",
             "ect_plo_lookup",
            "ect_gocprogrammingspvautomation_lookup",
            "ect_gocprogrammingspvfire_lookup",
            "ect_gocprogrammingspvsecurity_lookup",
            "ect_nocprogrammingspvautomation_lookup",
            "ect_nocprogrammingspvfire_lookup",
            "ect_nocprogrammingspvsecurity_lookup",
            "ect_ztcengineeringspvautomation_lookup",
            "ect_ztcengineeringspvfire_lookup",
            "ect_ztcengineeringspvsecurity_lookup"
             ));
      }
      /// <summary>
      /// This function will read the businessType.
      /// </summary>
      /// <param name="service">.</param>
      /// <param name="businessTypeNameId">.</param>
      /// <returns>.</returns>
      private static Entity ReadBusinessType(IOrganizationService service, Guid businessTypeNameId)
      {
         return service.Retrieve("ect_businesstype", businessTypeNameId, new ColumnSet("ect_businesstype"));
      }
      /// <summary>
      /// Read all the Project configuration.
      /// </summary>
      /// <param name="service">The service<see cref="IOrganizationService"/>.</param>
      /// <returns>The <see cref="EntityCollection"/>.</returns>
      private EntityCollection ReadProjectConfiguration(IOrganizationService service)
      {
         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='ect_configuration'>
                                <attribute name='ect_configurationid' />
                                <attribute name='ect_configurationname' />
                                <attribute name='createdon' />
                                <attribute name='ect_businesstypeid' />
                                <attribute name='ect_projectid' />
                                <attribute name='ect_minimumvalue_base' />
                                <attribute name='ect_minimumvalue' />
                                <attribute name='ect_maximumvalue_base' />
                                <attribute name='ect_maximumvalue' />
                                <attribute name='ect_configurationtype' />
                                <order attribute='ect_configurationname' descending='false' />
                                <filter type='and'>                             
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='ect_configurationtype' operator='eq' value='100000000' />
                                </filter>
                                </entity>
                        </fetch>";
         EntityCollection configuration = service.RetrieveMultiple(new FetchExpression(fetchString));
         return configuration;
      }
      /// <summary>
      /// Read all the Project configuration.
      /// </summary>
      /// <param name="service">The service<see cref="IOrganizationService"/>.</param>
      /// <returns>The <see cref="Entity"/>.</returns>
      private Entity ReadRFiConfiguration(IOrganizationService service)
      {
         Entity rfiConfiguration = null;

         string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='ect_configuration'>
                                <attribute name='ect_configurationname' />
                                <attribute name='createdon' />
                                <attribute name='ect_projectid' />
                                <attribute name='ect_minimumvalue' />
                                <attribute name='ect_maximumvalue' />
                                <attribute name='ect_configurationtype' />
                                <attribute name='ect_businesstypeid' />
                                <attribute name='ect_configurationid' />
                                <attribute name='ect_numberofdays_wholenumber' />
                                <order attribute='ect_configurationname' descending='false' />
                                <filter type='and'>
                                  <condition attribute='statecode' operator='eq' value='0' />
                                  <condition attribute='ect_configurationtype' operator='eq' value='402220002' />
                                </filter>
                              </entity>
                            </fetch>";

         EntityCollection rfiConfigurations = service.RetrieveMultiple(new FetchExpression(fetchString));

         if (rfiConfigurations.Entities.Count >= 1)
         {
            rfiConfiguration = rfiConfigurations.Entities[0];
         }
         return rfiConfiguration;
      }
   }
}
