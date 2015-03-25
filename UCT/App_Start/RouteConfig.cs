using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace UCT
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );

            routes.MapRoute(
                name: "ProgramLearningActivitiesCompetency",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "ProgramLearningActivitiesCompetency", action = "Index", id = UrlParameter.Optional }
             );

            routes.MapRoute(
                name: "Program",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Program", action = "Index", id = UrlParameter.Optional }
             );

           routes.MapRoute(
           name: "Competency",
           url: "{controller}/{action}/{id}",
           defaults: new { controller = "Competency", action = "Index", id = UrlParameter.Optional }
        );

          routes.MapRoute(
          name: "LearningActivities",
          url: "{controller}/{action}/{id}",
          defaults: new { controller = "LearningActivities", action = "Index", id = UrlParameter.Optional }
        );

          routes.MapRoute(
            name: "CompetencyLearningActivities",
            url: "{controller}/{action}/{id}",
            defaults: new { controller = "CompetencyLearningActivities", action = "Index", id = UrlParameter.Optional }
          );


          routes.MapRoute(
          name: "LearningActivitiesCompetencies",
          url: "{controller}/{action}/{id}",
          defaults: new { controller = "ProgramLearningActivities", action = "Index", id = UrlParameter.Optional }
        );

          routes.MapRoute(
          name: "ProgramCompetencies",
          url: "{controller}/{action}/{id}",
          defaults: new { controller = "ProgramCompetencies", action = "Index", id = UrlParameter.Optional }
        );
          routes.MapRoute(
            name: "Report",
            url: "{controller}/{action}/{id}",
            defaults: new { controller = "Report", action = "Index", id = UrlParameter.Optional }
          );

          routes.MapRoute(
            name: "Admin",
            url: "{controller}/{action}/{id}",
            defaults: new { controller = "Admin", action = "Index", id = UrlParameter.Optional }
            );

        }
    }
}