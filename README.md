# PIMRoles

This script attempts to provide a wrapper around the activation of PIM Roles.

It presents the user with the PIM Roles available to them to activate. They can select one
or more roles, provide a reason and duration they want the role activated for.

![image](https://github.com/user-attachments/assets/eeb05df7-fe2a-42e0-8362-f2ab44e4f294)

Clicking on Activate will activate all the selected Roles, with the Reason and Duration
input on the form.
The duration is subject to the maximum allowed for a Role, so the script will adjust the
duration for a role where this maximum is exceeded.

If a PIM Role is already activated it will be greyed out in the form.

![image](https://github.com/user-attachments/assets/a5b98fc2-f91e-4fa5-ba41-d4fd270a8c0b)

With every run of the script a History file is maintained at `"$env:USERPROFILE\Documents\PIMRoleSelections.json"`
