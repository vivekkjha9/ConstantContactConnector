using CTCT.Components;
using CTCT.Components.Contacts;
using CTCT.Exceptions;
using CTCT.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using SageCRMConnector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CTCT
{
    class CTCTMailListConnector1
    {
        static void MainMList(string[] args)
        {


            Uri OrganizationUri = new Uri(String.Format("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Organization.svc"));

            Uri HomeRealmUri = new Uri(String.Format("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Discovery.svc"));

            Uri serviceUri = new Uri("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Organization.svc");

            System.ServiceModel.Description.ClientCredentials clientCredentials = new System.ServiceModel.Description.ClientCredentials();

            clientCredentials.UserName.UserName = "crm@netunousa.com";
            clientCredentials.UserName.Password = "Busa2567";

            OrganizationServiceProxy serviceproxy = new OrganizationServiceProxy(OrganizationUri, null, clientCredentials, null);
            IOrganizationService orgservice = (IOrganizationService)serviceproxy;

            //ConditionExpression condition = new ConditionExpression();
            //condition.AttributeName = "new_sagecustomerid";
            //condition.Operator = ConditionOperator.Equal;
            //condition.Values.Add("0");
            string[] cols = { "listid", "listname", "new_constantcontactid","new_membercount"};
            ColumnSet columns = new ColumnSet(cols);
            QueryExpression expression = new QueryExpression();


            expression.EntityName = "list";
            //expression.ColumnSet = new AllColumns();
            expression.ColumnSet.AllColumns = true;

            QueryExpression query = new QueryExpression
            {
                EntityName = "list",
                ColumnSet = new ColumnSet("listid", "listname", "new_constantcontactid","new_membercount")
               
            };

            //expression.ColumnSet.AllColumns = true;
            expression.ColumnSet = columns;
            //expression.Criteria.AddCondition(condition);

            EntityCollection listEntityColl = orgservice.RetrieveMultiple(query);

            if (listEntityColl.Entities.Count > 0)
            {

                IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
                ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
                IListService listService = serviceFactory.CreateListService();
                for (int i = 0; i < listEntityColl.Entities.Count; i++)
                {


                    try
                    {



                        //get contact if exists (by email address)
                        CTCT.Components.Contacts.ContactList list = null;
                       

                        try
                        {
                            list = GetListByID(listEntityColl.Entities[i].Attributes["new_constantcontactid"].ToString());
                        }
                        catch (Exception ex)
                        {
                            CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                        }

                        bool alreadyExists = list != null ? true : false;

                        list = UpdateListFields(list, listEntityColl.Entities[i].Attributes);

                        CTCT.Components.Contacts.ContactList result = null;

                        // var contactService = _constantContactFactory.CreateContactService();

                        if (alreadyExists)
                        {
                            try
                            {
                                result = listService.UpdateList(list);
                                listEntityColl.Entities[i].Attributes["new_constantcontactid"] = result.Id;
                                listEntityColl.Entities[i].Attributes["new_membercount"] = list.ContactCount.ToString();
                                orgservice.Update((Entity)listEntityColl.Entities[i]);
                            }
                            catch(Exception ex)
                            {
                                CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                            }
                        }
                        else
                        {
                            try { 
                            result = listService.AddList(list);

                            listEntityColl.Entities[i].Attributes["new_constantcontactid"] = result.Id;
                            listEntityColl.Entities[i].Attributes["new_membercount"] = list.ContactCount.ToString();
                            orgservice.Update((Entity)listEntityColl.Entities[i]);
                            }
                            catch(Exception ex)
                            {
                                result = listService.UpdateList(list);

                                listEntityColl.Entities[i].Attributes["new_constantcontactid"] = result.Id;
                                listEntityColl.Entities[i].Attributes["new_membercount"] = list.ContactCount.ToString();
                                orgservice.Update((Entity)listEntityColl.Entities[i]);

                                CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());
                                alreadyExists = true;

                            }

                        }

                        if (result != null)
                        {
                            if (alreadyExists)
                            {
                                //messageResult = "Changes successfully saved!";


                            }
                            else
                            {
                                //  messageResult = "Contact successfully added!";
                            }
                        }
                        else
                        {
                            if (alreadyExists)
                            {
                                //    messageResult = "Failed to save changes!";
                            }
                            else
                            {
                                //      messageResult = "Failed to add contact!";
                            }
                        }

                        //MessageBox.Show(messageResult, "Result");





                    }
                    catch (IllegalArgumentException illegalEx)
                    {
                        //MessageBox.Show(GetExceptionsDetails(illegalEx, "IllegalArgumentException"), "Exception");
                        CTCTLogger.LogFile(illegalEx.InnerException.ToString(), illegalEx.InnerException.ToString(), illegalEx.Data.ToString(), (int)illegalEx.LineNumber(), illegalEx.Source.ToString());
                    }
                    catch (CtctException ctcEx)
                    {
                        //MessageBox.Show(GetExceptionsDetails(ctcEx, "CtctException"), "Exception");
                        CTCTLogger.LogFile(ctcEx.InnerException.ToString(), ctcEx.InnerException.ToString(), ctcEx.Data.ToString(), (int)ctcEx.LineNumber(), ctcEx.Source.ToString());
                    }
                    catch (OAuth2Exception oauthEx)
                    {
                        CTCTLogger.LogFile(oauthEx.InnerException.ToString(), oauthEx.InnerException.ToString(), oauthEx.Data.ToString(), (int)oauthEx.LineNumber(), oauthEx.Source.ToString());
                    }
                    catch (Exception ex)
                    {
                        CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                    }
                }

                //btnSave.Enabled = true;
            }
            // add all members in he list.

          
            
            PagingInfo pageInfo = new PagingInfo();
            pageInfo.Count = 5000;
            pageInfo.PageNumber = 1;


            
        }

        private static ContactList UpdateListFields(ContactList list, AttributeCollection attributeCollection)
        {
            if (list == null)
            {
                try { 
                CTCT.Components.Contacts.ContactList list1 = new ContactList();
                list1.Name = attributeCollection["listname"].ToString();
                list1.Status = Status.Active;
                list = list1;
                    }
                catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }
                
            }
            else
            {
                try
                {
                    list.Name = attributeCollection["listname"].ToString();

                    list.Status = Status.Active;
                }
                catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

            }
            return list;
        }

        private static CTCT.Components.Contacts.ContactList GetListByID(string p)
        {
            IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            IListService listService = serviceFactory.CreateListService();
            try { 
            ContactList lists = listService.GetList(p);

            if (lists != null)
            {
               
                    return lists;
                
               
            }
                }
          
                 catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());
                    return null;
                }
               
            

            return null;
            
        }

    }
}
