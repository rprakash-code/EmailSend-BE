Vendor Email Approval Workflow
This project facilitates the approval process for stakeholders to share documents with vendors via email. The workflow involves both frontend and backend components, ensuring seamless approval and secure document sharing.

How to Use
Stakeholders log in to the SharePoint site and attach the required document to a new request.
The system sends an approval email to the stakeholder's manager.
After manager approval, the backend system handles:
Generating a secure shared link for the document.
Sending the shared link to the specified vendors via email.


Workflow Overview
Frontend (SharePoint):

Stakeholders submit requests by uploading the necessary document to the SharePoint interface.
A request is created, and an approval email is triggered to the stakeholder's manager.
Approval Process:

Managers receive an email with details of the request and an option to approve or reject it.
On approval, the system proceeds to the next step.
Backend (Automated Email and Permissions):

The backend system generates a shared link to the document in SharePoint with the appropriate permissions.
The shared link is automatically sent to the designated vendors via email.



