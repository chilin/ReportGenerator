using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Client;
using System.Net;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ReportGeneratorApp.DataSource
{
    public class TFSHelper
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof (TFSHelper));
        Project project = null;
        string projectName = null;
        TfsTeamProjectCollection tfs = null;
        //readonly SqlConnection sqlCon = new SqlConnection(System.Configuration.ConfigurationManager.AppSettings["ConStr_Tfs_DefaultCollection"].ToString());

        public TFSHelper(string projectName)
        {
            this.projectName = projectName;
            GetWorkItemStore(projectName);
        }

        private void GetWorkItemStore(string projectName)
        {
            string uri = ConfigurationManager.AppSettings["TFSServer"];
            string userName = ConfigurationManager.AppSettings["UserName"];
            string password = ConfigurationManager.AppSettings["UserPw"];
            string domain = ConfigurationManager.AppSettings["Domain"];
            tfs = new TfsTeamProjectCollection(new Uri(uri), new NetworkCredential(userName, password, domain));
            try
            {
                tfs.EnsureAuthenticated();
            }
            catch (Exception ex)
            {
                log.Error("Connection TFS Server Error", ex);
                return;
            }
            if (tfs.HasAuthenticated)
            {
                try
                {
                    WorkItemStore Store = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));
                    if (Store.Projects.Contains(projectName))
                    {
                        project = Store.Projects[projectName];
                    }
                    else
                    {
                        log.Error("Get project Error. may be the project not exist");
                    }
                }
                catch (Exception ex)
                {
                    log.Error("Get Project Error", ex);
                }
            }
        }

        public Identity[] GetTfsGroups()
        {
            IGroupSecurityService gss = (IGroupSecurityService)tfs.GetService(typeof(IGroupSecurityService));
            ICommonStructureService css = (ICommonStructureService)tfs.GetService<ICommonStructureService>();
            return gss.ListApplicationGroups(css.GetProjectFromName(projectName).Uri);
        }

        public Identity[] GetMembers(params string[] groups)
        {
            IGroupSecurityService gss = (IGroupSecurityService)tfs.GetService(typeof(IGroupSecurityService));
            Identity[] tfsGroups = GetTfsGroups();
            List<Identity> members = new List<Identity>();
            List<Identity> tempMembers = new List<Identity>();
            foreach (string group in groups)
            {
                log.Info("Group Name : " + group);
                try
                {
                    if (string.IsNullOrWhiteSpace(group)) continue;
                    Identity tfsGroup = tfsGroups.FirstOrDefault(f => f.AccountName == group);
                    if (tfsGroup == null) continue;
                    Identity[] groupMembers = gss.ReadIdentities(SearchFactor.Sid, new string[] { tfsGroup.Sid }, QueryMembership.Expanded);
                    foreach (Identity member in groupMembers)
                    {
                        if (member.Members != null)
                        {
                            foreach (string memberSid in member.Members)
                            {
                                Identity memberInfo = gss.ReadIdentity(SearchFactor.Sid, memberSid, QueryMembership.None);
                                log.Info("AccountName : " + memberInfo.AccountName);
                                tempMembers.Add(memberInfo);
                            }
                        }
                    }
                    if (members.Count == 0)
                    {
                        log.Info("Add to email list");
                        members.AddRange(tempMembers);
                        tempMembers.Clear();
                    }
                    else
                    {
                        log.Info("Intersect email list");
                        List<Identity> temp = members.Intersect(tempMembers, new KeyEqualityComparer<Identity>(s => s.AccountName)).ToList();
                        log.Info("result list: " + string.Join(",", temp.Select(s => s.AccountName)));
                        members.Clear();
                        members.AddRange(temp);
                        log.Info("member list: " + string.Join(",", members.Select(s => s.AccountName)));
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }
            }
            return members.ToArray();
        }

        //public string[] GetMemberEmails(Identity[] members)
        //{
        //    List<string> emails = new List<string>();
        //    foreach (Identity member in members)
        //    {
        //        if (!string.IsNullOrWhiteSpace(member.MailAddress))
        //        {
        //            emails.Add(member.MailAddress);
        //        }
        //        else
        //        {
        //            emails.Add(Directory.GetEmailAddress(member.DisplayName));
        //        }
        //    }
        //    return emails.ToArray();
        //}

        public List<string> GetMemberList(params string[] groups)
        {
            Identity[] members = GetMembers(groups);
            List<string> names = new List<string>();
            foreach (Identity member in members)
            {
                names.Add(member.DisplayName);
            }
            return names;
        }
    }
}