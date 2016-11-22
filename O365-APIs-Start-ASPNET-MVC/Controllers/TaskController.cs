using System.Collections.Generic;
using System.Web.Mvc;

using O365_APIs_Start_ASPNET_MVC.Helpers;
using System.Threading.Tasks;
using model = O365_APIs_Start_ASPNET_MVC.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;


namespace O365_APIs_Start_ASPNET_MVC.Controllers
{
    [Authorize]
    public class TaskController : Controller
    {
        private static bool _O365ServiceOperationFailed = false;
        private EWSOperations _ewsOperation = new EWSOperations();
        // GET: Task
        public async Task<ActionResult> Index()
        {
            ViewBag.O365ServiceOperationFailed = _O365ServiceOperationFailed;

            if (_O365ServiceOperationFailed)
            {
                _O365ServiceOperationFailed = false;
            }

            List<model.EWSTaskItem> myTasks = new List<Models.EWSTaskItem>();
            try
            {
                myTasks = await _ewsOperation.getEWSTasks();
            }
            catch (AdalException e)
            {

                if (e.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {

                    //This exception is thrown when either you have a stale access token, or you attempted to access a resource that you don't have permissions to access.
                    throw e;

                }

            }

            return View(myTasks);
        }
    }
}