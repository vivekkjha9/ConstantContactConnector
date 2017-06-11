using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SageCRMConnector
{
    class Contact
    {




         try
                {
                    //get contact if exists (by email address)
                    Contact contact = null;

                    try
                    {
                       contact = GetContactByEmailAddress(txtEmail.Text.Trim());
                    }
                    catch (CtctException ctcEx)
                    {
                        //contact not found
                    }

                    bool alreadyExists = contact != null ? true : false;

                    contact = UpdateContactFields(contact);

                    Contact result = null;

                    var contactService = _constantContactFactory.CreateContactService();

                    if (alreadyExists)
                    {
                        result = contactService.UpdateContact(contact, false);
                    }
                    else
                    {
                        result = contactService.AddContact(contact, false);
                    }

                    if (result != null)
                    {
                        if (alreadyExists)
                        {
                            messageResult = "Changes successfully saved!";
                        }
                        else
                        {
                            messageResult = "Contact successfully added!";
                        }
                    }
                    else
                    {
                        if (alreadyExists)
                        {
                            messageResult = "Failed to save changes!";
                        }
                        else
                        {
                            messageResult = "Failed to add contact!";
                        }
                    }

                    MessageBox.Show(messageResult, "Result");
                }
                catch (IllegalArgumentException illegalEx)
                {
                    MessageBox.Show(GetExceptionsDetails(illegalEx, "IllegalArgumentException"), "Exception");
                }
                catch (CtctException ctcEx)
                {
                    MessageBox.Show(GetExceptionsDetails(ctcEx, "CtctException"), "Exception");
                }
                catch (OAuth2Exception oauthEx)
                {
                    MessageBox.Show(GetExceptionsDetails(oauthEx, "OAuth2Exception"), "Exception");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(GetExceptionsDetails(ex, "Exception"), "Exception");
                }

  
        }

       


        #region Private methods

        /// <summary>
        /// Get a contact by email address
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        private Contact GetContactByEmailAddress(string emailAddress)
        {
            var contactService = _constantContactFactory.CreateContactService();
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
        private Contact UpdateContactFields(Contact contact)
        {
            if (contact == null)
            {
                contact = new Contact();

                //add lists [Required]
                contact.Lists.Add(new ContactList() { Id = "1", Status = Status.Active });

                //add email_addresses [Required]
                var emailAddress = new EmailAddress() { 
                    Status = Status.Active, 
                    ConfirmStatus = ConfirmStatus.NoConfirmationRequired, 
                    EmailAddr = txtEmail.Text.Trim() 
                };
                contact.EmailAddresses.Add(emailAddress);   
            }

            contact.Status = Status.Active;

            #region Contact Information

            if (!string.IsNullOrWhiteSpace(txtFirstName.Text))
            {
                contact.FirstName = txtFirstName.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtMiddleName.Text))
            {
                contact.MiddleName = txtMiddleName.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtLastName.Text))
            {
                contact.LastName = txtLastName.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtHomePhone.Text))
            {
                contact.HomePhone = txtHomePhone.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtCity.Text) || !string.IsNullOrWhiteSpace(txtAddressLine1.Text) || 
                !string.IsNullOrWhiteSpace(txtAddressLine2.Text) || !string.IsNullOrWhiteSpace(txtAddressLine3.Text) ||
                !string.IsNullOrWhiteSpace(txtZip.Text) || !string.IsNullOrWhiteSpace(txtSubZip.Text) ||
                (cbCountry.SelectedItem != null))
            {
                Address address;

                if (contact.Addresses == null || contact.Addresses.Count() == 0)
                {
                    address = new Address();
                }
                else
                {
                    address = contact.Addresses[0];
                }

                if (!string.IsNullOrWhiteSpace(txtCity.Text))
                {
                    address.City = txtCity.Text.Trim();
                }

                if (!string.IsNullOrWhiteSpace(txtAddressLine1.Text))
                {
                    address.Line1 = txtAddressLine1.Text.Trim();
                }

                if (!string.IsNullOrWhiteSpace(txtAddressLine2.Text))
                {
                    address.Line2 = txtAddressLine2.Text.Trim();
                }

                if (!string.IsNullOrWhiteSpace(txtAddressLine3.Text))
                {
                    address.Line3 = txtAddressLine3.Text.Trim();
                }

                if (!string.IsNullOrWhiteSpace(txtZip.Text))
                {
                    address.PostalCode = txtZip.Text.Trim();
                }

                if (!string.IsNullOrWhiteSpace(txtSubZip.Text))
                {
                    address.SubPostalCode = txtSubZip.Text.Trim();
                }

                if (cbCountry.SelectedItem != null)
                {
                    CountryInfo selectedCountry = cbCountry.SelectedItem as CountryInfo;
                    if (selectedCountry != null)
                    {
                        address.CountryCode = selectedCountry.TwoLetterCountryName;
                    }
                }

                if (contact.Addresses.Count() == 0)
                {
                    contact.Addresses.Add(address);
                }
                else
                {
                    contact.Addresses[0] = address;
                }
            }

            #endregion Contact Information

            #region Company Information

            if (!string.IsNullOrWhiteSpace(txtCompanyName.Text))
            {
                contact.CompanyName = txtCompanyName.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtJobTitle.Text))
            {
                contact.JobTitle = txtJobTitle.Text.Trim();
            }

            if (!string.IsNullOrWhiteSpace(txtWorkPhone.Text))
            {
                contact.WorkPhone = txtWorkPhone.Text.Trim();
            }

            #endregion Company Information

            #region Notes

            if (!string.IsNullOrWhiteSpace(txtNotes.Text))
            {
                Note note = new Note();
                note.Content = txtNotes.Text.Trim();
                note.Id = "1";

                contact.Notes = new Note[] { note };
            }

            #endregion Notes

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

        #endregion Private methods

      
    }

    
}
    }
}
