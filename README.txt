Run the linked server script to add ADSI as a datasource in the form of a linked server. Replace CONTOSO/ADSERVICE with an account that has LDAP read privileges (hint most do by default-- don't put domain admin here dummy)

Run the stored procedure creation script to allow recursive group lookup. Replace dc=CONTOSO,dc=LOCAL with your domain.

Edit the report query as you see fit. You can build into a stored procedure or drop straight into a PowerBI/SSRS/etc report. Replace dc=CONTOSO,dc=LOCAL with your domain again here.