using System;
using System.ComponentModel.DataAnnotations;

namespace O365_APIs_Start_ASPNET_MVC.Models
{
    public class EWSTaskItem
    {
        public string taskId { get; set; }
        public string taskSubject { get; set; }
        public string taskStatus { get; set; }

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MM-yyyy}")]
        public DateTimeOffset? taskStartDate { get; set; }

        public EWSTaskItem(Microsoft.Exchange.WebServices.Data.Task exchanegTask)
        {
            //taskId = exchanegTask.Id.ToString();
            taskSubject = exchanegTask.Subject;
            taskStatus = exchanegTask.Status.ToString();
            if (exchanegTask.StartDate != null)
            { taskStartDate = exchanegTask.StartDate; }
            else { taskStartDate = null; }
        }

    }
}
