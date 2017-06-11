using CTCT.Components;
using CTCT.Components.Contacts;
using CTCT.Exceptions;
using CTCT.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using SageCRMConnector;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CTCT
{
    class CTCTMailListConnector
    {
        static void Main(string[] args)
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
            string[] cols = { "listid", "listname", "new_constantcontactid" };
            ColumnSet columns = new ColumnSet(cols);
            QueryExpression expression = new QueryExpression();


            expression.EntityName = "list";
            //expression.ColumnSet = new AllColumns();
            expression.ColumnSet.AllColumns = true;

            QueryExpression query = new QueryExpression
            {
                EntityName = "list",
                ColumnSet = new ColumnSet("listid", "listname", "new_constantcontactid"),

            };


            expression.ColumnSet = columns;


            EntityCollection listEntityColl = orgservice.RetrieveMultiple(query);
            IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            IListService listService = serviceFactory.CreateListService();
            if (listEntityColl.Entities.Count > 0)
            {

                //get contact if exists (by email address)
                CTCT.Components.Contacts.ContactList list = null;
              
                
                for (int i = 0; i < listEntityColl.Entities.Count; i++)
                {


                    try
                    {





                        try
                        {
                            list = GetListByID(listEntityColl.Entities[i].Attributes["new_constantcontactid"].ToString(),ref listService);
                        }
                        catch (Exception ex)
                        {
                            CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                        }
                        bool alreadyExists = list != null ? true : false;
                        try
                        {


                            list = UpdateListFields(list, listEntityColl.Entities[i].Attributes);


                        }
                        catch (Exception ex)
                        {
                            CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                        }
                        // var contactService = _constantContactFactory.CreateContactService();
                        CTCT.Components.Contacts.ContactList result = null;
                       // IListService listService = serviceFactory.CreateListService();
                        if (alreadyExists)
                        {
                            try
                            {
                               
                                result = listService.UpdateList(list);
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
                               // IBulkStatusService listService1 = serviceFactory.CreateBulkStatusService();

                                
                                result = listService.AddList(list);
                                listEntityColl.Entities[i].Attributes["new_constantcontactid"] = result.Id;
                                orgservice.Update((Entity)listEntityColl.Entities[i]);
                            }

                            catch (Exception ex)
                            {
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
                        CTCTLogger.LogFile(illegalEx.InnerException.ToString(), illegalEx.InnerException.ToString(), illegalEx.Data.ToString(), (int)illegalEx.LineNumber(), illegalEx.Source.ToString());
                    }
                    catch (CtctException ctcEx)
                    {

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
            

            if (listEntityColl.Entities.Count > 0)
            {

                bool flag = false; ;
              
                for (int i = 0; i < listEntityColl.Entities.Count; i++)
                {


                    try
                    {
                        GetAllMembersInaList(listEntityColl.Entities[i].Attributes["listid"].ToString(), listEntityColl.Entities[i].Attributes["new_constantcontactid"].ToString(), flag, ref orgservice,ref listService);
                    }
                    catch (Exception ex)
                    {
                        CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                    }
                    flag = true;
                }
            }



        }

        private static ContactList UpdateListFields(ContactList list, AttributeCollection attributeCollection)
        {
            if (list == null)
            {
                try
                {
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

        private static CTCT.Components.Contacts.ContactList GetListByID(string p, ref IListService listService)
        {
            //IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            //ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            //IListService listService = serviceFactory.CreateListService();
            try
            {
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

        private static void GetAllMembersInaList(string listid, string constantcontactlistid, bool flag, ref IOrganizationService orgservice, ref IListService listService)
        {
            ArrayList memberGuids = new ArrayList();
           
                    

            PagingInfo pageInfo = new PagingInfo();
            pageInfo.Count = 5000;
            pageInfo.PageNumber = 1;

            QueryByAttribute query = new QueryByAttribute("listmember");
            // pass the guid of the Static marketing list
            query.AddAttributeValue("listid", listid);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = orgservice.RetrieveMultiple(query);
            Entity entity1 = null;
            Contact result = null;
            IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            IContactService contactService = serviceFactory.CreateContactService();
           
            ContactList list = null;
            try
            {

                list = GetListByID(constantcontactlistid, ref listService);
            }
            catch (Exception ex)
            {
                CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

            }

            foreach (Entity entity in entityCollection.Entities)
            {
                try
                {
                    Contact contact = null;

                    entity1 = RetrieveEntityById(orgservice, "contact", Guid.Parse(((EntityReference)entity.Attributes["entityid"]).Id.ToString()));

                    if (entity1.Attributes.Contains("emailaddress1"))
                    {
                        contact = GetContactByEmailAddress(entity1.Attributes["emailaddress1"].ToString(), ref contactService);
                        if (flag == false)
                        {
                            if (contact != null)
                            {
                                contact.Lists.Clear();
                            }
                        }
                        bool alreadyExists = contact != null ? true : false;


                        try
                        {
                            contact = UpdateContactFields(contact, entity1.Attributes, list);
                        }
                        catch (Exception ex)
                        {
                            CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                        }


                        // var contactService = _constantContactFactory.CreateContactService();

                        if (alreadyExists)
                        {

                            try
                            {
                                result = contactService.UpdateContact(contact, false);
                                entity1.Attributes["new_integratectct"] = result.Id;
                                orgservice.Update((Entity)entity1);
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
                                result = contactService.AddContact(contact, false);
                                entity1.Attributes["new_integratectct"] = result.Id;
                                orgservice.Update((Entity)entity1);
                            }
                            catch (Exception ex)
                            {
                                CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

              
            }




            // if list contains more than 5000 records
            while (entityCollection.MoreRecords)
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = orgservice.RetrieveMultiple(query);



                foreach (Entity entity in entityCollection.Entities)
                {
                    try
                    {
                        memberGuids.Add(((EntityReference)entity.Attributes["entityid"]).Id);

                        entity1 = RetrieveEntityById(orgservice, "contact", Guid.Parse(((EntityReference)entity.Attributes["entityid"]).Id.ToString()));

                    }
                    catch (Exception ex)
                    {
                        CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                    }

                }
            }







        }



        private static Entity RetrieveEntityById(IOrganizationService service, string strEntityLogicalName, Guid guidEntityId)
        {
            Entity RetrievedEntityById = service.Retrieve(strEntityLogicalName, guidEntityId, new ColumnSet(true)); //it will retrieve the all attrributes
            return RetrievedEntityById;
        }

        private static Contact UpdateContactFields(Contact contact, AttributeCollection attr, ContactList ctctlistid)
        {
            try
            {
                if (contact == null)
                {
                    contact = new Contact();

                    //add lists [Required]
                    if (attr["donotbulkemail"].ToString().Equals("True"))
                    {
                        contact.Lists.Clear();
                        contact.Lists.Add(new ContactList() { Id = "2046326319", Status = Status.Active });
                        //  }

                    }

                    else
                    {
                        // contact.Lists.RemoveAll();

                        contact.Lists.Add(new ContactList() { Id = ctctlistid.Id.ToString(), Status = Status.Active });
                    }

                    //add email_addresses [Required]
                    var emailAddress = new EmailAddress()
                    {
                        Status = Status.Active,
                        ConfirmStatus = ConfirmStatus.NoConfirmationRequired,
                        EmailAddr = attr["emailaddress1"].ToString()
                    };
                    contact.EmailAddresses.Add(emailAddress);
                }

                contact.Status = Status.Active;

                #region Contact Information
                try
                {
                    if (attr["donotbulkemail"].ToString().Equals("True"))
                    {


                        contact.Lists.Clear();



                        contact.Lists.Add(new ContactList() { Id = "2046326319", Status = Status.Active });
                     

                    }
                    else
                    {


                        contact.Lists.Add(new ContactList() { Id = ctctlistid.Id.ToString(), Status = Status.Active });
                    }
                }
                catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

                if (!string.IsNullOrWhiteSpace(attr["firstname"].ToString()))
                {
                    contact.FirstName = attr["firstname"].ToString();
                }


                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["lastname"].ToString()))
                    {
                        contact.LastName = attr["lastname"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }
                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["mobilephone"].ToString()))
                    {
                        contact.HomePhone = attr["mobilephone"].ToString();
                    }
                }
                catch (Exception ex)
                {
                  //  CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["jobtitle"].ToString()))
                    {
                        contact.JobTitle = attr["jobtitle"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    //CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

                #endregion Contact Information

                #region Company Information

                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["company"].ToString()))
                    {
                        contact.CompanyName = attr["company"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    //CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }
                #endregion Company Information
                try
                {
                    //if (!string.IsNullOrWhiteSpace(attr["new_integratectct"].ToString()))
                    //{
                    //    if (contact.Id ==0)
                    //    { 
                    //      contact.Id = attr["new_integratectct"].ToString();
                    //    }
                    //}
                }
                catch (Exception ex)
                {
                    //CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["jobtitle"].ToString()))
                    {
                        contact.JobTitle = attr["jobtitle"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    //CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

                }


            }
            catch (Exception ex)
            {
                //CTCTLogger.LogFile(ex.InnerException.ToString(), ex.InnerException.ToString(), ex.Data.ToString(), (int)ex.LineNumber(), ex.Source.ToString());

            }
            return contact;
        }
        private static Contact GetContactByEmailAddress(string emailAddress, ref IContactService contactService)
        {
            //var contactService = _constantContactFactory.CreateContactService();
            // IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            // ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            // IContactService contactService = serviceFactory.CreateContactService();
            ResultSet<Contact> contacts = contactService.GetContacts(emailAddress, 1, null, null);
            try
            {

                if (contacts != null)
                {
                    if (contacts.Results != null && contacts.Results.Count > 0)
                    {
                        return contacts.Results[0];
                    }
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
