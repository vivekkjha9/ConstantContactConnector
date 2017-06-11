using CTCT.Components;
using CTCT.Components.Contacts;
using CTCT.Exceptions;
using CTCT.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CTCT
{
    class CTCTContactConnector
    {

               
           
         static void Mainss(string[] args)
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
            //condition.AttributeName = "new_integratectct";
            //condition.Operator = ConditionOperator.Equal;
            //condition.Values.Add("0");
            string[] cols = { "firstname", "lastname", "middlename", "emailaddress1", "address1_line1", "address1_line2", "address1_city", "address1_county", "address1_postalcode", "mobilephone", "donotbulkemail", "company", "jobtitle", "birthdate", "gendercode", "salutation" };
           ColumnSet columns = new ColumnSet(cols);
            QueryExpression expression = new QueryExpression();
            expression.EntityName = "contact";
            //expression.ColumnSet = new AllColumns();
            expression.ColumnSet.AllColumns = true;

            //expression.ColumnSet.AllColumns = true;
            expression.ColumnSet = columns;
            //expression.Criteria.AddCondition(condition);

            EntityCollection contactEntityColl = orgservice.RetrieveMultiple(expression);
           
            if (contactEntityColl.Entities.Count > 0)
            {

                IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
                ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
                IContactService contactService = serviceFactory.CreateContactService();
                for (int i = 0; i < contactEntityColl.Entities.Count; i++)
                {


         try
                {                  
        
        
        
                 //get contact if exists (by email address)
                    Contact contact = null;
                 if (contactEntityColl.Entities[i].Attributes.Contains("emailaddress1"))

                 { 
                    try
                    {
                        contact = GetContactByEmailAddress(contactEntityColl.Entities[i].Attributes["emailaddress1"].ToString());
                    }
                    catch (CtctException ctcEx)
                    {
                        //contact not found
                    }

                    bool alreadyExists = contact != null ? true : false;

                    contact = UpdateContactFields(contact, contactEntityColl.Entities[i].Attributes);

                    Contact result = null;

                   // var contactService = _constantContactFactory.CreateContactService();

                    if (alreadyExists)
                    {
                        result = contactService.UpdateContact(contact, false);
                        contactEntityColl.Entities[i].Attributes["new_integratectct"] = result.Id;
                        orgservice.Update((Entity)contactEntityColl.Entities[i]);

                    }
                    else
                    {
                        result = contactService.AddContact(contact, false);
                        contactEntityColl.Entities[i].Attributes["new_integratectct"] = result.Id;
                        orgservice.Update((Entity)contactEntityColl.Entities[i]);
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
                }
                catch (IllegalArgumentException illegalEx)
                {
                    //MessageBox.Show(GetExceptionsDetails(illegalEx, "IllegalArgumentException"), "Exception");
                }
                catch (CtctException ctcEx)
                {
                    //MessageBox.Show(GetExceptionsDetails(ctcEx, "CtctException"), "Exception");
                }
                catch (OAuth2Exception oauthEx)
                {
                    //MessageBox.Show(GetExceptionsDetails(oauthEx, "OAuth2Exception"), "Exception");
                }
                catch (Exception ex)
                {
                   // MessageBox.Show(GetExceptionsDetails(ex, "Exception"), "Exception");
                }
            }

            //btnSave.Enabled = true;
        }
         }

       
       

        #region Private methods

        /// <summary>
        /// Get a contact by email address
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        private static Contact GetContactByEmailAddress(string emailAddress)
        {
            //var contactService = _constantContactFactory.CreateContactService();
            IUserServiceContext userServiceContext = new UserServiceContext("3f09fe65-10ae-44a9-9db6-d9f2d6de1dec", "cvwkqk7ajrm67rcagvfwn9gx");
            ConstantContactFactory serviceFactory = new ConstantContactFactory(userServiceContext);
            IContactService contactService = serviceFactory.CreateContactService();
            ResultSet<Contact> contacts = contactService.GetContacts(emailAddress, 1, null, null);

            if (contacts != null)
            {
                if (contacts.Results != null && contacts.Results.Count > 0)
                {
                    return contacts.Results[0];
                }
            }

            return null;
        }

        /// <summary>
        ///Update contact fields based on completed form fields
        /// </summary>
        /// <param name="contact"></param>
        /// <returns></returns>
        private static Contact UpdateContactFields(Contact contact, AttributeCollection attr)
        {
            try
            {
                if (contact == null)
                {
                    contact = new Contact();

                    //add lists [Required]
                    if (attr["donotbulkemail"].ToString().Equals("True"))
                    {

                        contact.Lists.Add(new ContactList() { Id = "2046326319", Status = Status.Active });
                    }
                    //2046326319
                    else { 
                   
                        
                    contact.Lists.Add(new ContactList() { Id = "1005718013", Status = Status.Active });
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
                try {
                    if (attr["donotbulkemail"].ToString().Equals("True"))
                {
                    
                    contact.Lists.Add(new ContactList() { Id = "2046326319", Status = Status.Active });
                   
                }

                    }
                catch
                {


                }

                if (!string.IsNullOrWhiteSpace(attr["firstname"].ToString()))
                {
                    contact.FirstName = attr["firstname"].ToString();
                }

                //if (!string.IsNullOrWhiteSpace(attr["middlename"].ToString()))
                //{
                //    contact.MiddleName = attr["middlename"].ToString();
                //}

                if (!string.IsNullOrWhiteSpace(attr["lastname"].ToString()))
                {
                    contact.LastName = attr["lastname"].ToString();
                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["mobilephone"].ToString()))
                    {
                        contact.HomePhone = attr["mobilephone"].ToString();
                    }
                }
                catch (Exception)
                {
                    
                   // throw;
                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(attr["jobtitle"].ToString()))
                    {
                        contact.JobTitle = attr["jobtitle"].ToString();
                    }
                }
                catch (Exception)
                {

                    // throw;
                }

                

                //if (!string.IsNullOrWhiteSpace(attr["address_city1"].ToString()) || !string.IsNullOrWhiteSpace(attr["address_line1"].ToString()) ||
                //    !string.IsNullOrWhiteSpace(attr["address_line2"].ToString()) ||
                //    !string.IsNullOrWhiteSpace(attr["address_postalcode1"].ToString()))
                //{
                    //Address address;

                    //if (contact.Addresses == null || contact.Addresses.Count() == 0)
                    //{
                    //    address = new Address();
                    //}
                    //else
                    //{
                    //    address = contact.Addresses[0];
                    //}

                    
                    //    address.AddressType = ";
                   
                    //try
                    //{

                    //    if (!string.IsNullOrWhiteSpace(attr["address1_city1"].ToString()))
                    //    {
                    //        address.City = attr["address1_city1"].ToString();
                    //    }
                    //}
                    //catch (Exception)
                    //{
                        
                    //    throw;
                    //}
                    //try
                    //{

                    //    if (!string.IsNullOrWhiteSpace(attr["address1_line1"].ToString()))
                    //    {
                    //        address.Line1 = attr["address1_line1"].ToString();
                    //    }
                    //}
                    //catch (Exception)
                    //{
                        
                    //    throw;
                    //}

                    //try
                    //{
                    //    if (!string.IsNullOrWhiteSpace(attr["address1_line2"].ToString()))
                    //    {
                    //        address.Line2 = attr["address1_line2"].ToString();
                    //    }
                    //}
                    //catch (Exception)
                    //{
                        
                    //    throw;
                    //}


                    //try
                    //{

                    //    if (!string.IsNullOrWhiteSpace(attr["address1_postalcode"].ToString()))
                    //    {
                    //        address.PostalCode = attr["address1_postalcode"].ToString();
                    //    }
                    //}
                    //catch (Exception)
                    //{
                        
                        
                    //}



                    //if (contact.Addresses.Count() == 0)
                    //{
                    //    contact.Addresses.Add(address);
                    //}
                    //else
                    //{
                    //    contact.Addresses[0] = address;
                    //}
              //  }

                #endregion Contact Information

                #region Company Information

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(attr["company"].ToString()))
                        {
                            contact.CompanyName = attr["company"].ToString();
                        }
                    }
                    catch (Exception)
                    {
                        
                       // throw;
                    }

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(attr["jobtitle"].ToString()))
                        {
                            contact.JobTitle = attr["jobtitle"].ToString();
                        }
                    }
                    catch (Exception)
                    {
                        
                        //throw;
                    }

                //if (!string.IsNullOrWhiteSpace(txtWorkPhone.Text))
                //{
                //    contact.WorkPhone = txtWorkPhone.Text.Trim();
                //}

                #endregion Company Information

                //#region Notes

                //if (!string.IsNullOrWhiteSpace(txtNotes.Text))
                //{
                //    Note note = new Note();
                //    note.Content = txtNotes.Text.Trim();
                //    note.Id = "1";

                //    contact.Notes = new Note[] { note };
                //}

                //#endregion Notes
            }
            catch { 
            
            }
            return contact;
        }

        /// <summary>
        /// Get details for an exception
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="exceptionType"></param>
        /// <returns></returns>
        private string GetExceptionsDetails(Exception ex, string exceptionType)
        {
            StringBuilder sbExceptions = new StringBuilder();

            sbExceptions.Append(string.Format("{0} thrown:\n", exceptionType));
            sbExceptions.Append(string.Format("Error message: {0}", ex.Message));

            return sbExceptions.ToString();
        }



        private static void GetAllMembersInaList()
        {
            ArrayList memberGuids = new ArrayList();
            Uri OrganizationUri = new Uri(String.Format("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Organization.svc"));

            Uri HomeRealmUri = new Uri(String.Format("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Discovery.svc"));

            Uri serviceUri = new Uri("https://netunousa.api.crm.dynamics.com/XRMServices/2011/Organization.svc");

            System.ServiceModel.Description.ClientCredentials clientCredentials = new System.ServiceModel.Description.ClientCredentials();

            clientCredentials.UserName.UserName = "crm@netunousa.com";
            clientCredentials.UserName.Password = "Busa2567";

            OrganizationServiceProxy serviceproxy = new OrganizationServiceProxy(OrganizationUri, null, clientCredentials, null);
            IOrganizationService orgservice = (IOrganizationService)serviceproxy;

            PagingInfo pageInfo = new PagingInfo();
            pageInfo.Count = 5000;
            pageInfo.PageNumber = 1;

            QueryByAttribute query = new QueryByAttribute("listmember");
            // pass the guid of the Static marketing list
            query.AddAttributeValue("listid", new Guid("2CA7881F-3EDA-E111-B988-00155D886334"));
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = orgservice.RetrieveMultiple(query);

            foreach (Entity entity in entityCollection.Entities)
            {
                memberGuids.Add(((EntityReference)entity.Attributes["entityid"]).Id);
            }

            // if list contains more than 5000 records
            while (entityCollection.MoreRecords)
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = orgservice.RetrieveMultiple(query);

                foreach (Entity entity in entityCollection.Entities)
                {
                    memberGuids.Add(((EntityReference)entity.Attributes["entityid"]).Id);
                }
            }

        }

        #endregion Private methods

      
    }

    
}
    

